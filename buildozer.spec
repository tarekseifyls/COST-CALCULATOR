[app]
title = Cost Calculator
package.name = costcalc
package.domain = org.softech
source.dir = .
source.include_exts = py,png,jpg,kv,atlas,xlsx
version = 0.1

# 1. CLEAN REQUIREMENTS (Removed 'androidstorage4kivy' which causes crashes)
# Added 'pillow' to the list
requirements = python3,kivy,android,openpyxl,et_xmlfile,jdcal,pillow
orientation = portrait
fullscreen = 0

# 2. ANDROID CONFIG
android.api = 33
android.minapi = 21
android.accept_sdk_license = True

# 3. PERMISSIONS (The Master Key)
android.permissions = INTERNET,READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE,MANAGE_EXTERNAL_STORAGE

# 4. STABLE BRANCH
p4a.branch = master

[buildozer]
log_level = 2
warn_on_root = 0

