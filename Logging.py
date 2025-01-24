import os
import threading
from threading import Thread
import queue
import datetime  
from datetime import datetime
import time

#=============================================
#=============================================
#=============================================
#=============================================
# Tolog - renew log
def ToLog(message, startThread = False):
    global LogQueue
    try:
        LogQueue.put(str(datetime.today())[10:19] + "  " + str(message) + "\n")
        
    except NameError:
        print("creating LogQueue")
        LogQueue = queue.Queue()
        LogQueue.put(str(datetime.today())[10:19] + "  " + str(message) + "\n")
        
    except Exception as Err:
        print("Error in ToLog function, Error code = " + str(Err))
        
#=============================================
#=============================================
#=============================================
#=============================================
# Thread for saving logs
class LogThread(threading.Thread):
    def __init__(self, logdir):
        super().__init__()
        self.stop = False
        global LogQueue

        try:
            LogQueue
            
        except NameError:
            print("creating LgQueue")
            LogQueue = queue.Queue()
            
        self.logdir = logdir

    def run(self):
        ToLog("LogThread started!!!")
        self.writingQueue()
        ToLog("LogThread finished!!!")

    def writingQueue(self):
        global LogQueue
        while True:
            try:
                if LogQueue.empty():
                    if self.stop == True:
                        print("LogThreadStopped")
                        break
                    time.sleep(1)
                    continue
                else:
                    with open(self.logdir + "\\" + str(datetime.today())[0:10] + ".cfg", "a") as file:
                        while not LogQueue.empty():
                            mess = LogQueue.get_nowait()
                            file.write(mess)
                            #print("Wrote to Log:\t" + mess)
                        file.close()
            except Exception as Err:
                print("Error writing to Logfile, Error code = " + str(Err))
                #raise Exception
