from pyexcel.sheet import Sheet
from speedtest import Speedtest
import pyexcel as p
import bitmath
import os
import time

FILENAME = "speedtest.xlsx"
PAUSE_TIME = 60

client = Speedtest(config={"length" : {"download" : 30, "upload" : 30}})

def speedtest():
   print("Finding best server..")
   client.get_best_server()
   print("Testing Download...")
   client.download()
   print("Testing Upload...")
   client.upload()
   return client.results

def main():
   while True:
      results = speedtest()
      csv_res = results.csv().split(",")
      #csv_res = "4546,KEVAG Telekom GmbH,Koblenz,2020-10-08T21:17:44.102876Z,54.61139327239307,23.252,524642801.9837028,50559209.278758615,,5.146.199.169".split(",")
      csv_res[4] = int(float(csv_res[4]))
      csv_res[5] = int(float(csv_res[5]))
      csv_res[6] = int(bitmath.Bit(float(csv_res[6])).Mib)
      csv_res[7] = int(bitmath.Bit(float(csv_res[7])).Mib)

      if FILENAME not in os.listdir():
         sheet = Sheet(colnames=["Server ID","Sponsor","Server Name","Timestamp","Distance","Ping","Download","Upload","Share","IP Address"], sheet=[csv_res])
         sheet.save_as(FILENAME)
      else:
         sheet = p.get_sheet(file_name=FILENAME)
         sheet.extend_rows([csv_res])
         sheet.save_as(FILENAME)
      print("Saved to excel!")
      print(f"Waiting {PAUSE_TIME} Seconds")
      time.sleep(PAUSE_TIME)

if __name__ == "__main__":
   main() 