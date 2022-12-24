print('Loading Model...')
import multiprocessing
import cv2
import matplotlib.pyplot as plt
import os.path
import xlwt
from xlwt import Workbook
import xlrd
from xlutils.copy import copy
from datetime import datetime
from pathlib import Path
from playsound import playsound



def main_p():
    config_file = 'ssd_mobilenet_v3_large_coco_2020_01_14.pbtxt'
    frozen_model = 'frozen_inference_graph.pb'
    model = cv2.dnn_DetectionModel(frozen_model,config_file)
    model.setInputSize(320,320)
    model.setInputScale(1.0/127.5)
    model.setInputMean((127.5,127.5,127.5))
    model.setInputSwapRB(True)
    
    class_labels = []
    file_name = 'lable.txt'
    with open(file_name, 'rt') as fp:
        class_labels = fp.read().rstrip('\n').split('\n')
    
    #You can show all objects in lable or show specific objects
    print(" \n person , bicycle, motorbike, car, bus, train, truck \n")
      
    input_list=input("enter the objests to detect : ").split(' ')
    input_index=[]
    for x in input_list:
       y=class_labels.index(x)
       input_index.append(y)
    # print(input_index)
    
    print("1. enter video\n or\n2. open webcam")
    a=int(input("enter the choice : " ))
    if(a == 1):
      video_file_path = input('Enter file path : ')
      cap = cv2.VideoCapture(video_file_path)
    elif( a == 2):
        cap = cv2.VideoCapture(0)
    else:
        print("enter correct choice")

    if not cap.isOpened():
        raise IOError('Cannot open video')
    font = cv2.FONT_HERSHEY_PLAIN
    cap.set(cv2.CAP_PROP_FRAME_WIDTH,320)
    cap.set(cv2.CAP_PROP_FRAME_HEIGHT,320)
    
    while True:
        flag=False
        ret, frame = cap.read()
        class_index, confidence, bbox = model.detect(frame,confThreshold=0.5)
    
        if len(class_index) != 0:
            for class_ind, conf, boxes in zip(class_index.flatten(), confidence.flatten(), bbox):
                
                if class_ind-1 in input_index:
                    flag=True
                    cv2.rectangle(frame, boxes, (255,0,0), 2)
                    cv2.putText(frame, class_labels[class_ind-1],(boxes[0]+10,boxes[1]+40), font, fontScale=2, color=( 0,255, 0), thickness=2)
                    
                    
                    
        cv2.imshow('Object Detection', frame)
        if flag:
            alert(class_labels[class_ind-1])
        if cv2.waitKey(2) & 0xFF==ord('x'):
            break
    cap.release()
    cv2.destroyAllWindows()

alert_items=[]
def alert(x):
   
    now=datetime.now()
    now1=now.strftime("%B %d,/%Y %H:%M:%S")
  
    if x not in alert_items:
        path_to_file="records.xls"
        path=Path(path_to_file)
        if path.is_file():
            pass
        else:
           wb1=Workbook()
           s1=wb1.add_sheet("sheet 1")
           s1.write(0,0,"name")
           s1.write(0,1,"time")
           wb1.save("records.xls")
        rb=xlrd.open_workbook("records.xls")
        s2 =rb.sheet_by_index(0)                    
        s2.cell_value(0,0)                           
        rows= s2.nrows             
        cols= s2.ncols
        cols2=cols-cols
        cols3=cols2+1
        wb=copy(rb)
        w_sheet=wb.get_sheet(0)
        w_sheet.write(rows,cols2,x)
        w_sheet.write(rows,cols3,now1)
        wb.save("records.xls")                  
        alert_items.append(x)
        
        for i in range(2):
          playsound("sound file name")

main_p()
