set logfile="C:\Users\padwoadmin\Dropbox (Macmillan Publishers)\wordmaker_stg\stdout-and-err.txt"

cd S:\resources\bookmaker_scripts\wordmaker >> %logfile% 2>&1
git pull origin master >> %logfile% 2>&1