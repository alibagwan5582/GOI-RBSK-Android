import os, re, shutil
from datetime import datetime
from openpyxl import load_workbook
from kivy.app import App
from kivy.lang import Builder
from kivy.metrics import dp
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.label import Label
from kivy.uix.button import Button
from kivy.uix.spinner import Spinner
from kivy.uix.textinput import TextInput
from kivy.uix.scrollview import ScrollView
from kivy.uix.popup import Popup

# Android document picker imports are available only inside the APK.
try:
    from android import activity as android_activity
    from jnius import autoclass, jarray
    ANDROID_AVAILABLE = True
except Exception:
    android_activity = None
    autoclass = None
    jarray = None
    ANDROID_AVAILABLE = False

AGE_GROUPS = ["Select", "6 Weeks to 3 Years", "3 Years to 6 Years", "6 Years to 18 Years"]
SCREENING_TYPES = ["Select", "No. of children screened in current month", "Found positive for selected health conditions in current month/year", "Referred for this current month"]
REFERRALS = ["Select", "PHC", "RH", "SDH", "DH", "SNCU", "DEIC"]
MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
DISEASES = [
"Neural tube defect", "Down's Syndrome", "Cleft Lip & Palate", "Club foot", "Developmental dysplasia of the hip", "Congenital cataract", "Congenital deafness", "Congenital heart diseases", "Retinopathy of Prematurity", "Microcephaly", "Macrocephaly", "Severe Anaemia", "Vitamin A deficiency (Bitot spot)", "Vitamin D Deficiency, (Rickets)", "SAM", "Severe Thinning", "Obesity", "Goitre", "Severe Stunting", "Vitamin B complex deficiency", "Skin conditions", "Otitis media", "Rheumatic heart disease", "Reactive airway disease", "Dental conditions", "Convulsive disorders", "Childhood leprosy Disease", "Childhood T.B.", "Childhood Extra Pulmonary T. B.", "Vision impairment", "Hearing Impairment", "Neuro motor impairment", "Motor delay", "Cognitive delay", "Language delay", "Behaviour disorder (Autism)", "Learning disorder", "Attention deficit hyperactivity disorder", "Others", "Growing up concerns", "Substance abuse", "Feel depressed", "Delay in menstruation cycles", "Irregular periods", "Pain or burning sensation while urinating", "Discharge/ foul smelling discharge from the genitor-urinary area", "Pain during menstruation"]
FRAME1_MAP = {
"6 Weeks to 3 Years": {"screened":("C10","C11"),"cum":("D10","D11"),"positive":("F10","F11"),"referred":("G10","G11")},
"3 Years to 6 Years": {"screened":("I10","I11"),"cum":("J10","J11"),"positive":("L10","L11"),"referred":("M10","M11")},
"6 Years to 18 Years": {"screened":("O10","O11"),"cum":("P10","P11"),"positive":("R10","R11"),"referred":("S10","S11")}}
FRAME2_ROW_VALUES=[25,26,27,28,29,30,31,32,33,34,35,37,38,39,40,41,42,43,44,45,47,48,49,50,51,52,53,54,55,57,58,59,60,61,62,63,64,65,66,68,69,70,71,72,73,74,75]
FRAME3_ROW_VALUES=[78,80,82,84,86,88,90,92,94,96,98,100,102,104,106,108,110,112,114,116,118,120,122,124,126,128,130,132,134,136,138,140,142,144,146,148,150,152,154,156,158,160,162,164,166,168,170]
def norm(s):
    s=str(s or '').strip().lower().replace('\xa0',' ').replace('–','-').replace('—','-').replace('’',"'").replace('‘',"'")
    return re.sub(r'\s+',' ',s)
FRAME2_ROWS={norm(d):r for d,r in zip(DISEASES,FRAME2_ROW_VALUES)}
FRAME3_ROWS={norm(d):r for d,r in zip(DISEASES,FRAME3_ROW_VALUES)}
FRAME2_COLS={AGE_GROUPS[1]:('C','D'),AGE_GROUPS[2]:('F','G'),AGE_GROUPS[3]:('I','J')}
FRAME2_RESTRICTED={AGE_GROUPS[1]:{41,42,49,64,65,66,68,69,70,71,72,73,74,75},AGE_GROUPS[2]:{41,49,64,65,66,68,69,70,71,72,73,74,75},AGE_GROUPS[3]:{40,44}}
FRAME3_COLS={AGE_GROUPS[1]:{'PHC':'C','RH':'C','SDH':'D','DH':'D','SNCU':'D','DEIC':'E'},AGE_GROUPS[2]:{'PHC':'F','RH':'F','SDH':'G','DH':'G','SNCU':'G','DEIC':'H'},AGE_GROUPS[3]:{'PHC':'I','RH':'I','SDH':'J','DH':'J','SNCU':'J','DEIC':'K'}}
FRAME3_RESTRICTED={AGE_GROUPS[1]:{'Severe Thinning','Rheumatic heart disease'},AGE_GROUPS[2]:{'Severe Thinning','Rheumatic heart disease'},AGE_GROUPS[3]:{'SAM'}}

KV = '''
<RootUI>:
    orientation: 'vertical'
    padding: dp(10)
    spacing: dp(8)
    canvas.before:
        Color:
            rgba: (0.95, 0.97, 0.99, 1)
        Rectangle:
            pos: self.pos
            size: self.size
    Label:
        text: 'GOI FORMAT RBSK TEAM-1 LOHARA'
        size_hint_y: None
        height: dp(54)
        bold: True
        font_size: '18sp'
        color: 1,1,1,1
        canvas.before:
            Color:
                rgba: (0.09, 0.38, 0.75, 1)
            Rectangle:
                pos: self.pos
                size: self.size
    BoxLayout:
        size_hint_y: None
        height: dp(48)
        spacing: dp(5)
        Button:
            text: 'Target'
            on_release: root.show_page('target')
        Button:
            text: 'Screening'
            on_release: root.show_page('screening')
        Button:
            text: 'Referral'
            on_release: root.show_page('referral')
    BoxLayout:
        size_hint_y: None
        height: dp(42)
        Label:
            text: 'Reporting Month'
            size_hint_x: .42
            color: .1,.15,.2,1
        Spinner:
            id: month
            text: app.current_month
            values: app.months
            size_hint_x: .58
    ScrollView:
        do_scroll_x: False
        GridLayout:
            id: content
            cols: 1
            size_hint_y: None
            height: self.minimum_height
            spacing: dp(10)
            padding: dp(4)
    Label:
        id: status
        text: 'Select details and tap Save Data.'
        size_hint_y: None
        height: dp(42)
        color: .05,.35,.12,1
        text_size: self.width, None
'''

class RootUI(BoxLayout):
    page='target'
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.build_page('target')
    def show_page(self, page):
        self.page=page; self.build_page(page)
        self.ids.status.text=''
    def add_label(self, text):
        self.ids.content.add_widget(Label(text=text,size_hint_y=None,height=dp(32),color=(.08,.13,.2,1),halign='left',valign='middle',text_size=(self.width-dp(40),None)))
    def add_spinner(self, hint, values, height=46):
        sp=Spinner(text=hint,values=values,size_hint_y=None,height=dp(height))
        self.ids.content.add_widget(sp); return sp
    def add_input(self, hint, numeric=False):
        ti=TextInput(hint_text=hint,multiline=False,input_filter='int' if numeric else None,size_hint_y=None,height=dp(48),font_size='17sp')
        self.ids.content.add_widget(ti); return ti
    def build_page(self, page):
        box=self.ids.content; box.clear_widgets()
        self.add_label({'target':'TARGET ENTRY','screening':'SCREENING DETAILS','referral':'REFERRAL DETAILS'}[page])
        self.age=self.add_spinner('Select Age Group',AGE_GROUPS)
        if page=='target':
            self.criteria=self.add_spinner('Select Screening Type',SCREENING_TYPES)
        else:
            self.health=self.add_input('Search or type Health Condition')
            self.add_label('Health condition must match an item from the list below. Tap a condition to select it.')
            self.disease_spinner=self.add_spinner('Select Health Condition',DISEASES,60)
            self.health.bind(text=self.filter_diseases)
            self.disease_spinner.bind(text=self.choose_disease)
        if page=='referral':
            self.refer=self.add_spinner('Select Refer To',REFERRALS)
        self.male=self.add_input('Male count',True)
        self.female=self.add_input('Female count',True)
        action_row = BoxLayout(size_hint_y=None, height=dp(54), spacing=dp(8))
        btn=Button(text='SAVE DATA',size_hint_x=.55,bold=True,background_normal='',background_color=(.05,.52,.42,1))
        btn.bind(on_release=lambda *_: self.save_current())
        export_btn=Button(text='EXPORT EXCEL',size_hint_x=.45,bold=True,background_normal='',background_color=(.10,.38,.75,1))
        export_btn.bind(on_release=lambda *_: App.get_running_app().export_workbook())
        action_row.add_widget(btn); action_row.add_widget(export_btn)
        box.add_widget(action_row)
        if page!='target':
            note='Excel row mapping is preserved from the supplied GOI app.'
            self.add_label(note)
    def filter_diseases(self, instance, value):
        # Spinner's full values stay available; typing filters by selecting a matching option where possible.
        matches=[d for d in DISEASES if norm(value) in norm(d)] if value.strip() else DISEASES
        if matches:
            self.disease_spinner.values=matches
    def choose_disease(self, instance, value):
        if value and value!='Select Health Condition': self.health.text=value
    def save_current(self):
        try:
            age=self.age.text
            if age not in AGE_GROUPS[1:]: raise ValueError('Please select Age Group.')
            month=self.ids.month.text
            if month not in MONTHS: raise ValueError('Please select Reporting Month.')
            male=self.number(self.male.text,'Male'); female=self.number(self.female.text,'Female')
            app=App.get_running_app(); path=app.workbook_path
            if not os.path.isfile(path): raise FileNotFoundError('Workbook not found in app storage.')
            wb=load_workbook(path); ws=wb['Rural']
            ws['O5']=month
            if self.page=='target':
                criterion=self.criteria.text
                if criterion not in SCREENING_TYPES[1:]: raise ValueError('Please select Screening Type.')
                mapping=FRAME1_MAP[age]
                key={SCREENING_TYPES[1]:'screened',SCREENING_TYPES[2]:'positive',SCREENING_TYPES[3]:'referred'}[criterion]
                a,b=mapping[key]; ws[a]=male; ws[b]=female
                if key=='screened':
                    cm,cf=mapping['cum']; ws[cm]=self.number(ws[cm].value,'Cumulative Male')+male; ws[cf]=self.number(ws[cf].value,'Cumulative Female')+female
            else:
                health=self.health.text.strip()
                if not health: health=self.disease_spinner.text
                key=norm(health)
                if key not in FRAME2_ROWS: raise ValueError('Please select a valid Health Condition from the list.')
                if self.page=='screening':
                    row=FRAME2_ROWS[key]
                    if row in FRAME2_RESTRICTED.get(age,set()): raise ValueError(f'{health} is not applicable for selected Age Group.')
                    mc,fc=FRAME2_COLS[age]; ws[f'{mc}{row}']=male; ws[f'{fc}{row}']=female
                else:
                    if health not in FRAME3_RESTRICTED.get(age,set()) and key in FRAME3_ROWS:
                        row=FRAME3_ROWS[key]
                    else:
                        row=FRAME3_ROWS.get(key)
                    if row is None: raise ValueError('Invalid Health Condition.')
                    if health in FRAME3_RESTRICTED.get(age,set()): raise ValueError(f'{health} is not applicable for selected Age Group.')
                    refer=self.refer.text
                    if refer not in REFERRALS[1:]: raise ValueError('Please select Refer To.')
                    mc=FRAME3_COLS[age][refer]; fc=chr(ord(mc)+1); ws[f'{mc}{row}']=male; ws[f'{fc}{row}']=female
            wb.save(path); wb.close()
            self.ids.status.text='Saved successfully. Workbook updated in app storage.'
            self.male.text=''; self.female.text=''
        except Exception as e:
            self.ids.status.text='Error: '+str(e)
    @staticmethod
    def number(value,label):
        value=str(value or '').strip()
        if not value:return 0
        if not re.fullmatch(r'\d+',value): raise ValueError(f'{label} value must be a whole number.')
        return int(value)

class GOIRBSKApp(App):
    months=MONTHS
    current_month=datetime.now().strftime('%B')
    def build(self):
        self.title='GOI RBSK Team-1 Lohara'
        self.workbook_path=os.path.join(self.user_data_dir,'MHT_III_GOI_Team Wise_2026-27.xlsx')
        if not os.path.exists(self.workbook_path):
            os.makedirs(self.user_data_dir,exist_ok=True)
            bundled=os.path.join(os.path.dirname(__file__),'assets','MHT_III_GOI_Team Wise_2026-27.xlsx')
            shutil.copy2(bundled,self.workbook_path)
        if ANDROID_AVAILABLE:
            android_activity.bind(on_activity_result=self._on_export_result)
        Builder.load_string(KV)
        return RootUI()

    def export_workbook(self):
        # Android's system save dialog exports the latest workbook outside private app storage.
        if not os.path.isfile(self.workbook_path):
            self.root.ids.status.text = 'Export failed: Excel workbook not found.'
            return
        if not ANDROID_AVAILABLE:
            try:
                dest = os.path.join(os.path.expanduser('~'), 'GOI_RBSK_Export.xlsx')
                shutil.copy2(self.workbook_path, dest)
                self.root.ids.status.text = 'Excel exported to: ' + dest
            except Exception as exc:
                self.root.ids.status.text = 'Export failed: ' + str(exc)
            return
        try:
            Intent = autoclass('android.content.Intent')
            intent = Intent(Intent.ACTION_CREATE_DOCUMENT)
            intent.addCategory(Intent.CATEGORY_OPENABLE)
            intent.setType('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            intent.putExtra(Intent.EXTRA_TITLE, 'MHT_III_GOI_Team Wise_2026-27.xlsx')
            android_activity.startActivityForResult(intent, 7314)
            self.root.ids.status.text = 'Choose Downloads, Drive, or another folder to export Excel.'
        except Exception as exc:
            self.root.ids.status.text = 'Cannot open export picker: ' + str(exc)

    def _on_export_result(self, request_code, result_code, intent):
        if request_code != 7314:
            return
        try:
            Result = autoclass('android.app.Activity')
            if result_code != Result.RESULT_OK or intent is None:
                self.root.ids.status.text = 'Export cancelled.'
                return
            uri = intent.getData()
            Activity = autoclass('org.kivy.android.PythonActivity')
            stream = Activity.mActivity.getContentResolver().openOutputStream(uri)
            if stream is None:
                raise IOError('Android could not open the selected destination.')
            with open(self.workbook_path, 'rb') as source:
                data = source.read()
            for offset in range(0, len(data), 65536):
                stream.write(jarray('b')([b if b < 128 else b - 256 for b in data[offset:offset + 65536]]))
            stream.flush()
            stream.close()
            self.root.ids.status.text = 'Excel exported successfully. File is now outside the app.'
        except Exception as exc:
            self.root.ids.status.text = 'Export failed: ' + str(exc)

if __name__=='__main__': GOIRBSKApp().run()
