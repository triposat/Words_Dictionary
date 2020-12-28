from difflib import get_close_matches
import json
from plyer import notification
import time
book = json.load(open("ebooks.json"))
def speak(str):
 from win32com.client import Dispatch
 speak=Dispatch("SAPI.SpVoice")
 speak.Speak(str)
def teller(word):
    if word.lower() in book:
        return book[word]
    elif word.title() in book:
        return book[word.title()]
    elif word.upper() in book:
        return book[word.upper()]
    elif len(get_close_matches(word , book.keys())) > 0 :
        notification.notify(title=f"Showing results for {get_close_matches(word, book.keys())[0].capitalize()}",
                            message=f"Search instead for {word.capitalize()}",
                            app_icon="C:/Users"
                                     "/Dell/Downloads/sear.ico", timeout=5)
        speak(f"Showing results for {get_close_matches(word, book.keys())[0].capitalize()}")
        time.sleep(2)
        return book[get_close_matches(word,book.keys())[0]]
    else:
        notification.notify(title="Error!! Word not found in the Dictionary",
                            message="Please Try again",
                            app_icon="C:/Users"
                                     "/Dell/Downloads/err.ico", timeout=7)
        speak("Error!! Word not found in the Dictionary")
        speak("Please Try again")
        exit()


if __name__ == '__main__':
  notification.notify(title="Offline E-book Dictionary Made by : Satyam Tripathi",
                        message="Using Python & Python Plyer",
                        app_icon="C:/Users"
                                 "/Dell/Downloads/smile.ico", timeout=7)
  speak("Offline E-book Dictionary!! Made by!! Satyam Tripathi")
  time.sleep(1)
  while 1:
   word = input("Please Enter the Word : ")
   tran= teller(word)
   a=type(tran)
   if a==list:
    for i in tran:
        notification.notify(title=f"Meaning of {get_close_matches(word, book.keys())[0].capitalize()} : ",
                            message=f"{i}",
                            app_icon="C:/Users"
                                     "/Dell/Downloads/sear.ico", timeout=8)
        speak(f"Meaning of {get_close_matches(word, book.keys())[0].capitalize()}")
        speak(f"{i}")
        time.sleep(2)
        break
   else:
       speak(f"Meaning of {word.capitalize()}")
       speak(f"{tran}")
       notification.notify(title=f"Meaning of {word.capitalize()} : ",
                           message=f"{tran}",
                           app_icon="C:/Users"
                                    "/Dell/Downloads/sear.ico", timeout=8)
   a=input("Do you want Search another Words : Press[Y/N]").lower()
   if a=='y':
       continue
   else:
       notification.notify(title=f"Thanks a lot my dear",message="",
                           app_icon="C:/Users"
                                    "/Dell/Downloads/thank.ico", timeout=8)
       speak(f"Thanks a lot my dear ")
       speak("You are the best person in planet Earth\nHope You Will try again")
       exit()
