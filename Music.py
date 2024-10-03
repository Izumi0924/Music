# Encoding = UTF-8
import time
import os
import openpyxl
from pydub import AudioSegment
import subprocess
import json

def TimeCalculate(Time):
    Min, Last = Time.split(":")
    Sec, MilliSec = Last.split(".")
    return (int(Min) * 60 + int(Sec)) * 1000 + int(MilliSec)

class Music:
    def __init__(self, Name, LoopStart, LoopEnd, RhythmEnd):
        self.Name = Name
        self.LoopStart = TimeCalculate(LoopStart)
        self.LoopEnd = TimeCalculate(LoopEnd)
        self.RhythmEnd = TimeCalculate(RhythmEnd)

class SheetOperator:
    def __init__(self, SheetPath, RootPath):
        self.Sheet = None
        self.SheetPath = SheetPath
        self.RootPath = RootPath
        self.TargetRhythm = None
        self.LoopTime = 10
        self.PreludePath = PathFile["PreludePath"]
        self.LoopPath = PathFile["LoopPath"]
        self.EpisodePath = PathFile["EpisodePath"]
        self.Command = None
        self.PID = ''
        self.Volume = 0.5
        self.SystemCommand = ["help", "stop", "pause", "play", "find", "status", "repeat", "volume", "repeat_volume", "reload", "exit"]

        self.PreludeLength = 0
        self.LoopLength = 0
        self.EpisodeLength = 0

        self.IsPlaying = 0
        self.IsPaused = 0
        self.Music = None

        self.loadSheet()  # Creat Sheet Object

    def loadSheet(self):
        TimeSheet = {}  # A Dic mapping to Music class

        wb = openpyxl.load_workbook(self.SheetPath)
        sheet = wb["Sheet1"]

        for row in sheet[sheet.min_row + 1: sheet.max_row]:
            Rhythm = row[0].value.lower()
            if Rhythm not in TimeSheet:
                NewMusic = Music(row[1].value, row[2].value, row[3].value, row[4].value)
                TimeSheet[Rhythm] = NewMusic
        self.Sheet = TimeSheet
        print("Success to load TimeSheet")

    def commandReader(self):
        while self.Command == None:
            Parameter1 = None
            Parameter2 = None
            print("Please Input a Part of Rhythm Name & Looptime & volume, Or Input 'help' to Read Help, Or Input 'find' to Read Rhythm List")
            CmdArray = input().lower().strip().split(" ")

            if len(CmdArray) >= 2:  # Separate Command and Looptime and volume
                Parameter1 = CmdArray[1]
                if len(CmdArray) == 3:
                    Parameter2 = CmdArray[2]
                Cmd = CmdArray[0]
            else:
                Cmd = CmdArray[0]

            if Cmd == "":
                print("Command is Failed To Detect. Please Input Again")
                continue

            if Cmd in self.SystemCommand:
                self.Command = Cmd
                if self.Command == "repeat" or self.Command == "volume" or self.Command == "repeat_volume":
                    if not Parameter1 and not Parameter2:
                        self.Command = None
                        print("Parameter Missed. Please Input again")
                        continue
                if self.Command == "repeat":
                    Parameter1 = int(Parameter1)
                    if self.repeatParaCheck(Parameter1) == 1:
                        self.LoopTime = int(Parameter1)
                    else:
                        print("Looptime Input Error")
                        continue
                elif self.Command == "volume":
                    Parameter1 = float(Parameter1)
                    if self.volumeParaCheck(Parameter1) == 1:
                        self.Volume = Parameter1
                    else:
                        print("Volume Input Error")
                        continue
                elif self.Command == "repeat_volume":
                    Parameter1 = int(Parameter1)
                    Parameter2 = float(Parameter2)
                    if self.repeatParaCheck(Parameter1) == 1 and self.volumeParaCheck(Parameter2) == 1:
                        self.LoopTime = Parameter1
                        self.Volume = Parameter2
                    else:
                        print("Looptime or Volume Input Error")
                        continue
            else:
                CmdList = self.rhythmNameCheck(Cmd)
                if len(CmdList) == 0:
                    print("Failed to Match. Please Input Again")
                elif len(CmdList) == 1:
                    self.Command = CmdList[0]
                else:
                    self.Command = self.nameFilter(CmdList)
                try:
                    Parameter1 = int(Parameter1)
                    if self.repeatParaCheck(Parameter1) == 1:
                        self.LoopTime = Parameter1
                except:
                    self.LoopTime = 10
                try:
                    Parameter2 = float(Parameter2)
                    if self.volumeParaCheck(Parameter2) == 1:
                        self.Volume = Parameter2
                except:
                    self.Volume = 0.5

        # print(self.Command)

    def repeatParaCheck(self, Para):
        if Para < -1:
            return 0
        return 1

    def volumeParaCheck(self, Para):
        if Para < 0 or Para > 1:
            return 0
        return 1

    def commandExecutor(self):  # Main function
        if self.Command == "help":  # 查看命令帮助
            print("System Command:")
            print("    help  : Read help")
            print("    stop  : If there a music is playing or paused, stop it")
            print("    pause : If there a music is playing, pause it")
            print("    play  : If there a music is paused, resume it")
            print("    status: Read system status at the time")
            print("    repeat: Play existed Rhythm with new looptime")
            print("    volume: Change volume")
            print("    reload: Reload timesheet")
            print("repeat_volume: Change looptime and volume")

            print("    exit  : Exit this program")
            print("Input Format:")
            print("    A string witch is a prefix of a rhythm name + space + Looptime(Default 10) + space + Volume(Default 0.5)")
        elif self.Command == "stop":  # 停止当前音乐
            if self.IsPlaying == 1:
                self.PIDUpdate()
                if self.PID != "":
                    CommandStream = 'taskkill /pid ' + self.PID + ' /f'
                    os.system(CommandStream)
                self.IsPlaying = 0
                self.Music = None
                self.PIDFileUpdate()
            else:
                print("No Music is Playing")
        elif self.Command == "pause":  # 暂停当前音乐
            if self.IsPlaying == 1:
                CommandStream = "pssuspend64.exe" + " " + self.PID
                os.system(CommandStream)
                self.IsPaused = 1
                self.IsPlaying = 0
            else:
                print("No Music is Playing")
        elif self.Command == "play":  # 重启暂停的音乐
            if self.IsPaused == 1:
                CommandStream = "pssuspend64.exe" + " " + self.PID + " " + "-r"
                os.system(CommandStream)
                self.IsPaused = 0
                self.IsPlaying = 1
            else:
                print("No Music is Paused")
        elif self.Command == "find":  # 查看可播放音乐目录
            No = 1
            for key in self.Sheet.keys():
                print("Rhythm%d:%-10s" % (No, self.Sheet[key].Name))
                No += 1
        elif self.Command == "status":  # 查看当前运行状态
            self.PIDUpdate()
            if self.PID:
                if self.IsPaused == 1:
                    print("%-25s%s" % ("Rhythm is Pasued : ", self.Music))
                elif self.IsPlaying == 1:
                    print("%-25s%s" % ("Rhythm is Playing : ", self.Music))
                    print("%-25s%s" % ("Looptime is : ", self.LoopTime))
                    print("%-25s%s" % ("Volume is : ", self.Volume))
                print("%-25s%s" % ("Process ID is : ", self.PID))
            else:
                print("No Rhythm is Playing")
        elif self.Command == "repeat" or self.Command == "volume" or self.Command == "repeat_volume":
            if self.Music != None:  # 仅在有音乐文件的时候才可以执行 即打开程序不能直接执行该命令
                if self.IsPlaying == 1 or self.IsPaused == 1:
                    self.PIDUpdate()
                    if self.PID != "":
                        CommandStream = 'taskkill /pid ' + self.PID + ' /f'
                        os.system(CommandStream)
                self.playRhythm()
                self.IsPlaying = 1
                self.IsPaused = 0
            else:
                print("No Rhythm to Play")
        elif self.Command == "reload":  # 重新读入timesheet 一般不用
            self.loadSheet()
        elif self.Command == "exit":  # 退出程序
            self.PIDUpdate()
            if self.PID != "":
                CommandStream = 'taskkill /pid ' + self.PID + ' /f'
                os.system(CommandStream)
                self.PIDFileUpdate()

            time.sleep(0.1)  # 为了保证文件删除是在播放进程被kill掉之后 需要有一个较小的延时
            if self.Music != None:  # 仅在存在音乐时删除
                os.remove(self.PreludePath)
                os.remove(self.LoopPath)
                os.remove(self.EpisodePath)

            PID = os.getpid()  # 获取的PID是int
            CommandStream = 'taskkill /pid ' + str(PID) + ' /f'
            os.system(CommandStream)
        else:  # 播放音乐
            if self.IsPlaying == 1 or self.IsPaused == 1:  # 如果正在播放音乐 则先把正在播放的音乐停掉
                self.PIDUpdate()
                if self.PID != "":
                    CommandStream = 'taskkill /pid ' + self.PID + ' /f'
                    os.system(CommandStream)
            self.writeRhythm(self.Command)
            self.playRhythm()
            self.IsPlaying = 1
            self.Music = self.Sheet[self.Command].Name

        self.Command = None

    def rhythmNameCheck(self, Part):
        List = []
        for key, value in self.Sheet.items():
            if key.startswith(Part):
                List.append(key)
        return List

    def nameFilter(self, NameList):
        NewList = NameList
        while len(NewList) == 0 or len(NewList) > 1:
            print("Input a Aspect to Select a Rhythm From : ")
            No = 1
            for item in NewList:
                print("Rhythm%d : %s" % (No, item))
                No += 1
            NewList = []
            NewPart = input().lower()
            while NewPart == "":
                print("Pleasr Input Again")
                NewPart = input().lower()
            for Name in NameList:
                if NewPart in Name:
                    NewList.append(Name)
                    break
        return NewList[0]

    def writeRhythm(self, RhythmName):
        print("Rhythm Selected : ",self.Sheet[RhythmName].Name)
        RhythmPath = os.path.join(self.RootPath, self.Sheet[RhythmName].Name)

        Rhythm = AudioSegment.from_mp3(RhythmPath)
        Prelude = Rhythm[: self.Sheet[RhythmName].LoopEnd]
        Loop = Rhythm[self.Sheet[RhythmName].LoopStart: self.Sheet[RhythmName].LoopEnd]
        Epilogue = Rhythm[self.Sheet[RhythmName].LoopStart:]

        self.PreludeLength = self.Sheet[RhythmName].LoopEnd / 1000
        self.LoopLength = (self.Sheet[RhythmName].LoopEnd - self.Sheet[RhythmName].LoopStart) / 1000
        self.EpisodeLength = (self.Sheet[RhythmName].RhythmEnd - self.Sheet[RhythmName].LoopStart) / 1000

        Prelude.export(self.PreludePath, format='mp3')
        Loop.export(self.LoopPath, format='mp3')
        Epilogue.export(self.EpisodePath, format='mp3')
        # print("Rhythm output success")

    def playRhythm(self):
        print("Ready to play")
        subprocess.Popen(["python", "RhythmPlay.py", str(self.LoopTime), str(self.Volume), str(self.PreludeLength), str(self.LoopLength), str(self.EpisodeLength)], shell=True)
        print("python RhythmPlay.py", str(self.LoopTime), str(self.Volume), str(self.PreludeLength), str(self.LoopLength), str(self.EpisodeLength))
        time.sleep(5)   # 5秒内保护 不允许有任何操作
        self.PIDUpdate()

    def PIDUpdate(self):
        File = open("E:/python/Music/PID.txt", "r")
        self.PID = File.readline()
        File.close()

    def PIDFileUpdate(self):
        file = open("E:/python/Music/PID.txt", "w")
        file.close()

    def FreeProcess(self):
        self.PIDUpdate()
        if self.PID != "":
            CommandStream = 'taskkill /pid ' + self.PID + ' /f'
            os.system(CommandStream)

if __name__ == '__main__':
    with open("E:/python/Music/Path.json") as text:
        PathFile = json.load(text)

    SheetPath = PathFile["SheetPath"]
    RootPath = PathFile["RootPath"]

    RhythmSheet = SheetOperator(SheetPath, RootPath)

    RhythmSheet.PIDFileUpdate()
    while True:
        RhythmSheet.commandReader()
        RhythmSheet.commandExecutor()


