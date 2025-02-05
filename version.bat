@echo off

for /f "tokens=2 delims==" %%a in ('wmic datafile where "name='C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe'" get Version /value') do echo %%a > chrome_version.data

for /f "tokens=2 delims==" %%a in ('wmic datafile where "name='c:\\program files (x86)\\Microsoft\\Edge\\Application\\msedge.exe'" get Version /value') do echo %%a > edge_version.data
