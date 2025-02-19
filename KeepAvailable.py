import pyautogui
import math
import time

def circle():
    a,b = pyautogui.position()
    w = 20
    m = (2*math.pi)/w
    r = 200      

    while 1:
        time.sleep(30)
        for i in range(w+1):
            x = int(a+r*math.sin(m*i))  
            y = int(b+r*math.cos(m*i))
            pyautogui.moveTo(x, y, duration = 0.2)
        pyautogui.press("shift")
        print("Cursor and Keyboard are activated!")

if __name__ == "__main__":
    try:
        circle()
    except KeyboardInterrupt as e:
        print("Keep up the good work... ;-)")