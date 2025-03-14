#importing the neccessary modules

import win32com.client as w
import speech_recognition as sr
import webbrowser
import os 
import time
import pygame
import keyboard
import requests
#text-to-speech converting function

def say(text):
    speak=w.Dispatch("SAPI.SpVoice")
    speak.rate=0
    speak.Speak(text)

#speech recognition function

def mic():
    say("listening...")
    print()
    print("listening")
    r=sr.Recognizer()
    try:
        with sr.Microphone() as source:
            audio=r.listen(source)
            say("recognizing")
            print("recognizing...")
            query=r.recognize_google(audio,language="en-in")
            print(f"user said:{query}")
            print()
            text=query
            say(text)
            return text
           
    except Exception as e:
        print(f"Error: {e}")

#weather function

#music library

def music_library():
    print("=======================Music-Library==================================")
    say("welcome to music library section")
    print()
    say("you can access your favourite songs listed here")
    print("1.dont let me down(play 1)")
    say("dont let me down")
    print("2.into your arms(play 2")
    say("into your arms")
    print("3.shape of you(play 3)")
    say("shape of you")
    print("4.chiki miki othhe jhhara(play 4)")
    say("chiki miki othhe jhhara")
    print("5.sadhhaba bohu")
    say("sadhhaba bohu")
    print()
    songs=[["play don't let me down",r"C:\Users\swain\OneDrive\Desktop\spotify\musicfolder\Don't-Let-Me-Down(PagalNew.Com.Se).mp3"],
            ['play two',r"C:\Users\swain\OneDrive\Desktop\spotify\musicfolder\Into-Your-Arms(PagalNew.Com.Se).mp3"],
            ['play three',r"C:\Users\swain\OneDrive\Desktop\spotify\musicfolder\Shape-Of-You(PagalNew.Com.Se).mp3"],
            ['play four',r"C:\Users\swain\OneDrive\Desktop\spotify\musicfolder\Chiki-miki-othe-jhara-(Dharma-Nikiti)-(OdiaFresh.Com).mp3"],
            ['play five',r"C:\Users\swain\OneDrive\Desktop\spotify\musicfolder\Sadhaba-Bohu-Lo-(OdiaFresh.Com).mp3"]
    ]
    pygame.mixer.init()  # Initialize pygame mixer

    is_loop = True
    while is_loop:
        try:
            print("\nNeo: Which song do you want to play?")
            say("Which song do you want to play?")
            text = mic()

            if text:
                played = False
                for song in songs:
                    if song[0].lower() in text.lower():
                        print(f"Neo: Playing {song[0]}")
                        say(f"Playing {song[0]}")
                        played = True

                        # Play song using pygame
                        pygame.mixer.music.load(song[1])
                        pygame.mixer.music.play()

                        print("Press 'Enter' to stop the song...")
                        while pygame.mixer.music.get_busy():
                            if keyboard.is_pressed("enter"):
                                print("Neo: Stopping the song")
                                say("Stopping the song")
                                pygame.mixer.music.stop()
                                break
                            time.sleep(0.1)

                if not played:
                    print("Neo: Sorry! That song is not in the list.")
                    say("Sorry! That song is not available right now.")

            print("\nNeo: Do you want to return to the main menu? (Yes/No)")
            say("Do you want to return to the main menu? Say yes to go back, or no to stay here.")
            allow = mic()
            if allow and "yes" in allow.lower():
                say("Returning to the main menu")
                print("Neo: Returning to the main menu...")
                is_loop = False
            else:
                say("You chose to stay in the music library.")
        except Exception:
            say("Sorry, I couldn't recognize you. Say again.")
            print("Could not recognize you! Try again.")
#web browser

def web_browser():
    print("=======================Web Browser==================================")
    say("welcome to web browser section")
    print()
    say("you can access these few applications")
    print("1.Youtube")
    say("youtube")
    print("2.google")
    say("google")
    print("3.instagram")
    say("instagram")
    print()
    browsers=[['youtube','https://www.youtube.com'],
            ['google','https://www.google.com'],
            ['instagram','https://www.instagram.com']
    ]
    isloop=True
    while(isloop):
        try:    
            print("neo said: what do you want from the web browser section")
            say("so tell me what do you want")
            text=mic()
            opened=0
            for website in browsers:
                
                if f'{website[0]}'.lower() in text.lower():
                    print(f"neo said: opening {website[0]} for you sir")
                    say(f"opening {website[0]} sir ")
                    webbrowser.open (website[1])
                    opened=opened+1 
            if opened==0:
                print("neo : it is not in the option")
                say("sorry!it is not available right now")
            else:
                opened=0
            say("do you want to return to the main menu !type yes for returning to the main menu!or no for statying here")
            allow=input("neo said: do you want to return to the main menu(y/n)")
            if f'{allow}'.lower()=='y':
                    say("you choosed to go back to the main menu")
                    print("neo said :going back to main menu")
                    print("=====================main menu====================")
                    print("1.computer-internal-application")
                    print("2.web_browser")
                    print()
                    say("choose from this")
                    isloop=False
                    break   
            else:
                    isloop=True
                    say("you choosed to stay here in web browser section")
        except Exception:
            say("sorry ! i couldnot recognize you.say again")
            print(" ")
#computer internal application

def computer_internal_application():
    print("=======================computer-interal-application==================================")
    say("welcome to computer internal application")
    print()
    say("you can access these few applications")
    print("1.vlc media player")
    say("vlc mediaplayer")
    print("2.vs code")
    say("vs code")
    print("3.whatsapp")
    say("whatsapp")
    print()
    applications=[["vlc media",r"C:\Users\swain\OneDrive\Desktop\drive"],
                  ["vs code",r"C:\Users\swain\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Visual Studio Code\Visual Studio Code.lnk"],
                  ["my folder",r"C:\Users\swain\OneDrive\Desktop\pruthi"]
                  ]
    isloop=True
    while(isloop):
        try:    
            print("neo said: what do you want from the computer internal application")
            say("so tell me what do you want")
            text=mic()
            opened=0
            for app in applications:
                
                if f'{app[0]}'.lower() in text.lower():
                    print(f"neo said: opening {app[0]} for you sir")
                    say(f"opening {app[0]} sir ")
                    os.startfile(f"{app[1]}")
                    opened=opened+1 
            if opened==0:
                print("neo : it is not in the option")
                say("sorry!it is not available right now")
            else:
                opened=0
            say("do you want to return to the main menu !type yes for returning to the main menu!or no for statying here")
            allow=input("neo said: do you want to return to the main menu(y/n)")
            if f'{allow}'.lower()=='y':
                    say("you choosed to go back to the main menu")
                    print("neo said :going back to main menu")
                    print("=====================main menu====================")
                    print("1.computer-internal aplication")
                    print("access the web browser")
                    print()
                    say("choose from this")
                    isloop=False
                    break   
            else:
                    isloop=True
                    say("you choosed to stay here in computer internal application")
        except Exception:
            say("sorry ! i couldnot recognize you.say again")
            print(" ")
     
#final code
if __name__=="__main__":
    print("==============================================WELCOME:VOICE ASSISTANT=========================================================")
    say("welcome!TO THE voice assistant here you can acess some application by the voice command!hii i am neo and i am your guide")
    print("NOTE:\n--> i am  just a voice assistant which will guide you for some specific tasks.\n--> i will give you options for the task which  can be performed\n-->i will guide you so plz kindly ask me from the option\n-->remember i am not an  ai so i can do only those things if that would be listed. ")
    print()
    #main menu
    print("=========================main menu=============================")
    say("your main menu is here")
    say("your options are here")
    print("1.computer applications")
    say("computer appllications")
    print("2. Access Web browser")
    say("web browser")
    print("4.music library")
    say("music library")
    print("5.weather report(city of india)")
    say("weather report")
    print()
    loop=True
    while(loop):
        
        print("neo: now choose from the options")
        say("choose from the above option")
        try:    
            text=mic()
            if 'computer application'.lower() in text.lower():
                computer_internal_application()
                loop=True
            elif 'browser'.lower() in text.lower():
                web_browser()
                loop=True
            elif 'music'.lower() in text.lower():
                music_library()
                loop=True
            elif 'weather'.lower() in text.lower():
                weather_report()
                loop=True
            elif 'exit'.lower() in text.lower():
                say("exiting sir!see you again")
                loop=False
            else:
                print("neo: sorry!it is not in the option")
                say("it is not in option!try again")
        except Exception as e:
            print("couldnot recognize you!try again")
            say("sorry!try again")
            print(" ")
