import os
import sys
from pygame import mixer
import pygame
import time
import json
from PidManager import PidManager

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
    pidManager = PidManager()

    mixer.init()
    pygame.init()
    LoopTime = int(sys.argv[1])
    volume = float(sys.argv[2])
    # 记录PID
    pidManager.RecordPid(str(os.getpid()))

    PreludeLength = float(sys.argv[3])
    LoopLength = float(sys.argv[4])
    EpisodeLength = float(sys.argv[5])

    pygame.mixer.music.set_volume(volume)
    play(PreludeLength, PathFIle["PreludePath"], 1)
    play(LoopLength, PathFIle["LoopPath"], LoopTime)
    play(EpisodeLength, PathFIle["EpisodePath"], 1,)
    # 播放完成后，要清空PID
    pidManager.CleanPidFile()