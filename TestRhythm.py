import os
import sys
from pygame import mixer
import pygame
import time
import json

def play(Sleeptime, Path, LoopTime):
    mixer.music.load(Path)
    mixer.music.play(loops=LoopTime)
    if LoopTime == -1:
        while True:
            pass
    else:
        time.sleep(Sleeptime * LoopTime)

if __name__ == "__main__":
    with open("E:/python/Music/Path.json") as text:
        PathFIle = json.load(text)
    mixer.init()
    pygame.init()
    LoopTime = int(sys.argv[1])
    volume = float(sys.argv[2])

    file = open(PathFIle["PIDPath"], "w")  # 获取PID并输出
    file.write(str(os.getpid()))
    file.close()

    LoopLength = float(sys.argv[3])
    pygame.mixer.music.set_volume(volume)
    play(LoopLength, PathFIle["LoopPath"], LoopTime)

    file = open(PathFIle["PIDPath"], "w")  # 获取PID并输出
    file.close()