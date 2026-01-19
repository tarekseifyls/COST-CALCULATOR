[app]
title = Cost Calculator
package.name = costcalc
package.domain = org.softech

source.dir = .
source.include_exts = py,png,jpg,kv,atlas,xlsx

version = 0.1

requirements = kivy,openpyxl,androidstorage4kivy,et_xmlfile,jdcal

orientation = portrait
fullscreen = 0

android.api = 33
android.minapi = 21
android.accept_sdk_license = True
android.permissions = READ_MEDIA_DOCUMENTS


[buildozer]
log_level = 2
warn_on_root = 0


