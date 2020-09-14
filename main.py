#Author:  -S-TeaM
#Website: https://teamvietdev.com/
#Name: -S-chedules
#
#
import os
import time
start_time = time.time()

#run here
print('====================================')
print ('||Welcome to -S-TeaM with UET-VNU ||')
print('====================================')

relax = input('Nghỉ trưa hay ko: (y/n)')
huser = input("You are a real HUSer??(y/n) ")

#Frame
if relax == 'y' or relax == 'Y' :
    exec(open('frameRelaxLunch.py').read())
else:
    exec(open('frame.py').read())
#Copy Raw file

exec(open('copyRaw.py').read())
#Process main
exec(open('pretreatment.py').read())

if huser == 'y' or huser == 'Y' :
    exec(open('processingKHTNS.py').read())
else :
    exec(open('processing.py').read())
exec(open('processing.py').read())

exec(open('post_processing.py').read())
#open demo
print ('Nhớ kéo khoá quần!!!!')

os.system("out.xlsx")
#Check end
#print ('===Successfull main.py===')
