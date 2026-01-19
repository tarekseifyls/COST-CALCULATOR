[app]
title = Cost Calculator
package.name = costcalc
package.domain = org.softech
source.dir = .
source.include_exts = py,png,jpg,kv,atlas,xlsx
version = 0.1
requirements = python3,kivy,openpyxl,android

# Permissions
android.permissions = READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE

# Android Config
android.api = 33
android.minapi = 21
android.accept_sdk_license = True
orientation = portrait
fullscreen = 0

[buildozer]
log_level = 2
warn_on_root = 0