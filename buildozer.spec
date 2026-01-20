[app]
title = Cost Calculator
package.name = costcalc
package.domain = org.softech
source.dir = .
source.include_exts = py,png,jpg,kv,atlas,xlsx
version = 0.1

# We need these exact requirements for Excel to work on Android
requirements = python3,kivy,android,openpyxl,et_xmlfile,jdcal

orientation = portrait
fullscreen = 0

android.api = 33
android.minapi = 21
android.accept_sdk_license = True

# *** THE FIX: ADD MANAGE_EXTERNAL_STORAGE ***
android.permissions = INTERNET,READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE,MANAGE_EXTERNAL_STORAGE

# Branch configuration
p4a.branch = develop

[buildozer]
log_level = 2
warn_on_root = 0
