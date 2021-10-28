import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from win32com.client import Dispatch


def speak(str):
    speak = Dispatch(("SAPI.SpVoice"))
    speak.Speak(str)

if __name__ == '__main__':
    speak("hi saii welcome to ip project")
    speak("good morning examiner,iam deepaa, i will be assisting in the project made by saii")
    speak("the project is mainly based on the economic impact due to covid-19 pandemic")

x=np.arange(12)
w=["Global","India","China","Netherlands","Spain","US","UK","France","Euro Area","Germany","Japan","Italy"]
a={'No_Corona_Virus':[3,5.5,5.5,1.5,1.5,1.5,1,1.2,1.2,0.8,0.5,0.1],
    'Base_Line_Scenario':[1.5,5.2,2.2,0.8,0.8,1,0.5,0.5,-0.1,-0.5,-0.6,-1.8],
    'Risk_scenario_Pandemic':[0.8,3.5,0.1,-0.2,-0.5,-0.6,-0.8,-1,-1,-1.2,-2.5,-2]}


dataframe = pd.DataFrame(a,index=w)
dataframe.to_csv("S:\\GitHub\\12-python-projectip\\project.csv")

speak("the data frame has been saved to the location")

No_Corona_Virus=[3,5.5,5.5,1.5,1.5,1.5,1,1.2,1.2,0.8,0.5,0.1]
Base_Line_Scenario=[1.5,5.2,2.2,0.8,0.8,1,0.5,0.5,-0.1,-0.5,-0.6,-1.8]
Risk_scenario=[0.8,3.5,0.1,-0.2,-0.5,-0.6,-0.8,-1,-1,-1.2,-2.5,-2]

def mainmenu():
    choice = 0
    while choice != 4:
        print("\n---------------------------------------------")
        print("     GLOBAL ECONOMIC IMPACT OF COVID-19")
        print("\n---------------------------------------------")
        print("\n       1.DISPLAY WHOLE DATA")
        print("\n       2.DISPLAY GRAPHS")
        print("\n       3.DATA ANALYSIS")
        print("\n       4.EXIT MENU")
        print("\n---------------------------------------------")
        if __name__ == '__main__':
            speak("type a number for the required content")
        choice = int(input("Enter required value here: "))
        
        if choice == 1:
          if __name__ == '__main__':
            speak("this is the required data set")
          print(dataframe)
          
        elif choice == 2:
          subchoice = 0
          while subchoice !=3:
            if __name__ == '__main__':
              speak("type a number for which type of graph to be displayed")
              print("\n---------------------------------------------")
              print("\n       1.DISPLAY BAR GRAPH")
              print("\n       2.DISPLAY LINE GRAPH")
              print("\n       3.EXIT THE GRAPHS SHOWING MENU")
              print("\n---------------------------------------------")
            subchoice = int(input("enter required value here: "))
            if subchoice == 1:
              if __name__ == '__main__':
                speak("this is the required bar graph")
              plt.title("GLOBAL ECONOMIC IMPACT OF COVID-19",color="C70039",fontsize=24)
              plt.xticks(x,w,rotation=30)
              plt.bar(x - 0.25,No_Corona_Virus,color="blue",width=0.25,label="No Corona Virus")
              plt.bar(x,Base_Line_Scenario,color="red",width=0.25,label="Base Line Scenario (with Corona Virus)")
              plt.bar(x + 0.25,Risk_scenario,color="green",width=0.25,label="Risk scenario:Pandemic")
              plt.ylabel("ECONOMIC GROWTH(%)")
              plt.legend()
              plt.show()
              
            elif subchoice == 2:
              if __name__ == '__main__':
                speak("this is the required line graph")
              plt.title("GLOBAL ECONOMIC IMPACT OF COVID-19",color="C70039",fontsize=24)
              plt.grid(True)
              plt.xticks(x,w,rotation=30)
              plt.plot(x,No_Corona_Virus,color="blue",label="No Corona Virus")
              plt.plot(x,Base_Line_Scenario,color="red",label="Base Line Scenario (with Corona Virus)")
              plt.plot(x,Risk_scenario,color="green",label="Risk scenario:Pandemic")
              plt.ylabel("ECONOMIC GROWTH(%)")
              plt.legend()
              plt.show()
        elif choice == 3:
          subchoice1 = 0
          while subchoice1 !=5:
            if __name__ == '__main__':
              speak("type a number for the required data analysis")
            print("\n------------------------------------------------------------")
            print("\n       1.DISPLAY MAX VALUE OF ECONOMIC GROWTH COUNTRY")
            print("\n       2.DISPLAY MIN VALUE OF ECONOMIC GROWTH COUNTRY")
            print("\n       3.DISPLAY THE REQUIRED COLUMN")
            print("\n       4.DISPLAY THE REQUIRED ROW")
            print("\n       5.EXIT DATA ANALYSIS")
            print("\n------------------------------------------------------------")
            subchoice1 = int(input("enter required value here: "))
            if subchoice1 == 1:
                print(dataframe.loc["India",:])
                if __name__ == '__main__':
                    speak("India has the maximum economy in covid 19 crisis")
            if subchoice1 == 2:
                print(dataframe.loc["Italy",:])
                if __name__ == '__main__':
                    speak("Italy has the minimum economy in covid 19 crisis")     
            if subchoice1 == 3:
              subsubchoice = 0
              while subsubchoice !=4:
                print("\n---------------------------------------------")
                print("\n       1.DISPLAY NO CORONA VIRUS COLUMN")
                print("\n       2.DISPLAY BASE LINE SCENARIO COLUMN")
                print("\n       3.DISPLAY RISK SCENARIO COLUMN")
                print("\n       4.EXIT COLUMNS")
                print("\n---------------------------------------------")
                if __name__ == '__main__':
                    speak("type a number for the required column")
                subsubchoice = int(input("enter required value here: "))
                if subsubchoice == 1:
                  print(dataframe.No_Corona_Virus)
                  if __name__ == '__main__':
                    speak("this is the required column")
                elif subsubchoice == 2:
                  print(dataframe.Base_Line_Scenario)
                  if __name__ == '__main__':
                    speak("this is the required column")
                elif subsubchoice == 3:
                  print(dataframe.Risk_scenario_Pandemic)
                  if __name__ == '__main__':
                    speak("this is the required column")
                    
            if subchoice1 == 4:
              subsubchoice1 = 0
              while subsubchoice1 !=13:
                print("\n---------------------------------------------")
                print("\n       1.DISPLAY GLOBAL ROW")
                print("\n       2.DISPLAY INDIA ROW")
                print("\n       3.DISPLAY CHINA ROW")
                print("\n       4.DISPLAY NETHERLANDS ROW")
                print("\n       5.DISPLAY SPAIN ROW")
                print("\n       6.DISPLAY UNITED STATES ROW")
                print("\n       7.DISPLAY UNITED KINGDOM ROW")
                print("\n       8.DISPLAY FRANCE ROW")
                print("\n       9.DISPLAY EURO ASIA ROW")
                print("\n       10.DISPLAY GERMANY ROW")
                print("\n       11.DISPLAY JAPAN ROW")
                print("\n       12.DISPLAY ITALY ROW")
                print("\n       13.EXIT ROWS")
                print("\n---------------------------------------------")
                if __name__ == '__main__':
                    speak("type a number for the required row")
                subsubchoice1 = int(input("Enter required value here: "))
                if subsubchoice1 == 1:
                  print(dataframe.loc["Global",:])
                  if __name__ == '__main__':
                    speak("this is the required ROW")
                elif subsubchoice1 == 2:
                  print(dataframe.loc["India",:])
                  if __name__ == '__main__':
                    speak("this is the required ROW")
                elif subsubchoice1 == 3:
                  print(dataframe.loc["China",:])
                  if __name__ == '__main__':
                    speak("this is the required ROW")
                elif subsubchoice1 == 4:
                  print(dataframe.loc["Netherlands",:])
                  if __name__ == '__main__':
                    speak("this is the required ROW")
                elif subsubchoice1 == 5:
                  print(dataframe.loc["Spain",:])
                  if __name__ == '__main__':
                    speak("this is the required ROW")
                elif subsubchoice1 == 6:
                  print(dataframe.loc["US",:])
                  if __name__ == '__main__':
                    speak("this is the required ROW")
                elif subsubchoice1 == 7:
                  print(dataframe.loc["UK",:])
                  if __name__ == '__main__':
                    speak("this is the required ROW")
                elif subsubchoice1 == 8:
                  print(dataframe.loc["France",:])
                  if __name__ == '__main__':
                    speak("this is the required ROW")
                elif subsubchoice1 == 9:
                  print(dataframe.loc["Euro Area",:])
                  if __name__ == '__main__':
                    speak("this is the required ROW")
                elif subsubchoice1 == 10:
                  print(dataframe.loc["Germany",:])
                  if __name__ == '__main__':
                    speak("this is the required ROW")
                elif subsubchoice1 == 11:
                  print(dataframe.loc["Japan",:])
                  if __name__ == '__main__':
                    speak("this is the required ROW")
                elif subsubchoice1 == 12:
                  print(dataframe.loc["Italy",:])
                  if __name__ == '__main__':
                    speak("this is the required ROW")


mainmenu()


print("\n-------------------------------")
print("\n     THANK YOU FOR WATCHING")
print("\n-------------------------------")

if __name__ == '__main__':
            speak("thank you for watching, feel free to watch again")


