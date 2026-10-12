[app]
title = GOI RBSK Team-1 Lohara
package.name = goirbsk
package.domain = org.rbsk.lohara
source.dir = .
source.include_exts = py,kv,png,jpg,atlas,xlsx,txt
source.include_patterns = assets/*
version = 1.0.1
requirements = python3,kivy,openpyxl,pyjnius
orientation = portrait
fullscreen = 0
android.api = 35
android.minapi = 23
android.archs = arm64-v8a, armeabi-v7a
android.accept_sdk_license = True

[buildozer]
log_level = 2
warn_on_root = 1
