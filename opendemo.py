#open demo code by multicore
import os
import time
import multiprocessing 
def multiprocessing_func(x):
    os.system("G:\Code\schedules\out.xlsx")
    os.system("timeout /t 5 /nobreak && C:\Windows\System32\\taskkill.exe /IM EXCEL.EXE /F")
    
if __name__ == '__main__':
    starttime = time.time()
    processes = []
    for i in range(0,3):
        p = multiprocessing.Process(target=multiprocessing_func, args=(i,))
        processes.append(p)
        p.start()
        
    for process in processes:
        process.join()

#Check end
#print ('===Successfull opendemo.py===')
