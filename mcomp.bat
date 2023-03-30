@echo off
cls
echo '[GSR Compression and Archive creation Batch File]
echo '[-----------------------------------------------]

compress l_master.rpt l_master.rp_
compress l_roster.rpt l_roster.rp_
compress l_stfdet.rpt l_stfdet.rp_
compress l_stftim.rpt l_stftim.rp_
compress p_detstf.rpt p_detstf.rp_
compress p_except.rpt p_except.rp_
compress p_genstf.rpt p_genstf.rp_
compress p_rosdet.rpt p_rosdet.rp_
compress p_roster.rpt p_roster.rp_
compress p_stfdet.rpt p_stfdet.rp_
compress p_stftim.rpt p_stftim.rp_
compress p_master.rpt p_master.rp_

compress gsr.exe gsr.ex_
compress gsr.hlp gsr.hl_
REM compress gsr.dat gsr.da_

compress readme.txt readme.tx_
compress register.txt register.tx_

copy readme.txt website\GenSR\download
copy register.txt website\GenSR\download
copy gsr.hlp website\GenSR\download
copy *.rpt website\GenSR\download
copy gsr.exe website\GenSR\download
copy gsr.exe \gsr
copy gsr.hlp \gsr
copy *.rpt \gsr

move *.??_ website\GenSR\download

cd website\GenSR\download
pkzip -f gsr3031.zip
pkzip -f gsr3035x.zip
pkzip -f gsr3031f.zip
del *.??_
del gsr.hlp
del gsr.exe
del *.rpt
cd \contract\gsr
echo '[COMPLETE]

