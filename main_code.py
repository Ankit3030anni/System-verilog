import os
from threading import Thread

#reading all the block path
block_path = input ("block's path is : ")
a = input("python script path is : ");
file_name = open(f'{block_path}','r');
Lines = file_name.readlines();
def main_task(line):
  try:
    if(len(line.strip().split()) > 1):
      l1 = line.strip().split()[-1];
      
      os.system (f"python {a} {l1}")
  except:
    print("OK");


for line in Lines:
  #main_task(line);
  Thread(target=main_task, args =(line,)).start();


file_name.close();
