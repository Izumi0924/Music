# Encoding = UTF-8
import time
import os
import openpyxl
from pydub import AudioSegment
import subprocess
import json
import abc
from PidManager import PidManager

COMMAND_SUCC = 0
COMMAND_FAILED = 1

DEFAULT_LOOPTIME = 10
DEFAULT_VOLUME = 0.05

def TimeCalculate(Time):
    Min, Last = Time.split(":")
    Sec, MilliSec = Last.split(".")
    return (int(Min) * 60 + int(Sec)) * 1000 + int(MilliSec)

class Music:
    def __init__(self, name, loopStart, loopEnd, rhythmEnd):
        self.name = name
        self.loopStart = TimeCalculate(loopStart)
        self.loopEnd = TimeCalculate(loopEnd)
        self.rhythmEnd = TimeCalculate(rhythmEnd)

class MusicInfo:
    def __init__(self, name, looptime, volume):
        self.name = name
        self.looptime = looptime
        self.volume = volume

class CommandChecker():
    def __init__(self):
        pass

    def CheckVolume(self, volume) -> int:
        if (volume < 0 or volume > 1):
            return COMMAND_FAILED
        return COMMAND_SUCC

    def CheckLooptime(self, looptime) -> int:
        if (looptime == 0 or looptime < -1):
            return COMMAND_FAILED
        return COMMAND_SUCC

class RhythmStatus:
    def __init__(self):
        self.rhythm = None

        self.isPlaying = 0
        self.isPaused = 0

        self.volume = 0.05
        self.loopTime = 10

class RhythmCommander:
    # timeSheet: Dict(str, Music)
    # pathFile: Dict(str, str)
    # rhythmStatus: RhythmStatus
    # command: List[]
    def __init__(self, timeSheet, pathFile, rhythmStatus, command):
        self.timeSheet = timeSheet
        self.pathFile = pathFile
        self.rhythmStatus = rhythmStatus
        self.newRhythmStatus = rhythmStatus
        self.command = command

        self.commandChecker = CommandChecker()
        self.pidManager = PidManager()
    
    def _PlayRhythm(self, rhythmStatus):
        print("Ready to play")
        loopTIme = rhythmStatus.loopTime
        volume = rhythmStatus.volume

        ret = self._GetRhythmLength(rhythmStatus.rhythm)
        if (ret == COMMAND_FAILED):
            print("PLAY ERROR: Failed to get rhythm information")
            return
        preludeLength = ret[1][0]
        loopLength = ret[1][1]
        episodeLength = ret[1][2]
        subprocess.Popen(["python", "RhythmPlay.py", str(loopTIme), str(volume), str(preludeLength), str(loopLength), str(episodeLength)], shell=True)
        print("python RhythmPlay.py", str(loopTIme), str(volume), str(preludeLength), str(loopLength), str(episodeLength))
        time.sleep(5)

    def _KillOldRhythm(self):
        pid = self.pidManager.GetPid()
        if pid != "":
            try:
                commandStream = 'taskkill /pid ' + pid + ' /f'
                os.system(commandStream)
            except:
                pass

    def _GetRhythmLength(self, rhythmName) -> tuple():
        try:
            thisRhythmInfo = self.timeSheet[rhythmName]
            preludeLength = thisRhythmInfo.loopEnd / 1000
            loopLength = (thisRhythmInfo.loopEnd - thisRhythmInfo.loopStart) / 1000
            episodeLength = (thisRhythmInfo.rhythmEnd - thisRhythmInfo.loopStart) / 1000
            return (COMMAND_SUCC, [preludeLength, loopLength, episodeLength])
        except:
            return (COMMAND_FAILED, [" "])


    @abc.abstractmethod
    def OperateCommand(self) -> tuple():
        pass

class ShowHelp(RhythmCommander):
    def OperateCommand(self) -> tuple():
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

        return (COMMAND_SUCC, self.newRhythmStatus)

class StopRhythm(RhythmCommander):

    def __StopRhythm(self, pid):
        if pid != "":
            commandStream = 'taskkill /pid ' + pid + ' /f'
            os.system(commandStream)
    
    def UpdateNewRhythmStatus(self):
        self.newRhythmStatus.isPlaying = 0
        self.newRhythmStatus.rhythm = None
    
    def OperateCommand(self) -> tuple():
        if (self.rhythmStatus.isPlaying == 1):
            pid = self.pidManager.GetPid()
            self.__StopRhythm(pid)
            self.UpdateNewRhythmStatus()
            self.pidManager.CleanPidFile()
        else:
            print("No Music is Playing")
        
        return (COMMAND_SUCC, self.newRhythmStatus)

class PauseRhythm(RhythmCommander):
    
    def UpdateNewRhythmStatus(self):
        self.newRhythmStatus.isPaused = 1
        self.newRhythmStatus.isPlaying = 0
    
    def OperateCommand(self) -> tuple():
        if self.rhythmStatus.isPlaying == 1:
            pid = self.pidManager.GetPid()
            commandStream = "pssuspend64.exe" + " " + pid
            os.system(commandStream)
            self.UpdateNewRhythmStatus()
        else:
            print("No Music is Playing")
        return (COMMAND_SUCC, self.newRhythmStatus)

class RestartRhythm(RhythmCommander):
    def UpdateNewRhythmStatus(self):
        self.newRhythmStatus.isPlaying = 1
        self.newRhythmStatus.isPaused = 0

    def OperateCommand(self) -> tuple():
        if self.rhythmStatus.isPaused == 1:
            pid = self.pidManager.GetPid()
            commandStream = "pssuspend64.exe" + " " + pid + " " + "-r"
            os.system(commandStream)
            self.UpdateNewRhythmStatus()
        else:
            print("No Music is Paused")
        return (COMMAND_SUCC, self.newRhythmStatus)

class ShowAllRhythm(RhythmCommander):
    def OperateCommand(self) -> tuple():
        No = 1
        for key in self.timeSheet.keys():
            print("Rhythm%d:%-10s" % (No, self.timeSheet[key].name))
            No += 1

        return (COMMAND_SUCC, self.newRhythmStatus)

class ShowStatus(RhythmCommander):
    def OperateCommand(self) -> tuple():
        pid = self.pidManager.GetPid()
        if pid != "":
            if self.rhythmStatus.isPaused == 1:
                print("%-25s%s" % ("Rhythm is Pasued : ", self.rhythmStatus.rhythm))
            elif self.IsPlaying == 1:
                print("%-25s%s" % ("Rhythm is Playing : ", self.rhythmStatus.rhythm))
                print("%-25s%s" % ("Looptime is : ", self.rhythmStatus.loopTime))
                print("%-25s%s" % ("Volume is : ", self.rhythmStatus.volume))
            print("%-25s%s" % ("Process ID is : ", pid))
        else:
            print("No Rhythm is Playing")
        return (COMMAND_SUCC, self.newRhythmStatus)

class ChangeLooptime(RhythmCommander):
    def __GetTargetLooptime(self) -> tuple():
        if (len(self.command) != 2):
            print("Parameter Missed. Please Input again")
            return (COMMAND_FAILED, 0)
        try:
            looptime = int(self.command[1])
            if self.commandChecker.CheckLooptime(looptime) == COMMAND_SUCC:
                return (COMMAND_SUCC, looptime)
            print("Looptime Input Error")
            return (COMMAND_FAILED, 0)
        except:
            print("Looptime Input Error")
            return (COMMAND_FAILED, 0)

    def __UpdateLoopTime(self, targetLoopTime):
        self.newRhythmStatus.loopTime = targetLoopTime

    def __UpdateRhythmStatus(self):
        self.newRhythmStatus.isPlaying = 1
        self.newRhythmStatus.isPaused = 0

    def OperateCommand(self) -> tuple():
        if self.rhythmStatus.isPlaying == 0 and self.rhythmStatus.isPaused == 1:
            print("No Rhythm to Play")
            return (COMMAND_FAILED, self.rhythmStatus)
        
        ret = self.__GetTargetLooptime()
        if (ret[0] == COMMAND_FAILED):
            return (COMMAND_FAILED, self.rhythmStatus)
        self.__UpdateLoopTime(ret[1])
        self._KillOldRhythm()

        self._PlayRhythm(self.newRhythmStatus)
        self.__UpdateRhythmStatus()
        return (COMMAND_SUCC, self.newRhythmStatus)

class ChangeVolume(RhythmCommander):
    def __GetTargetVolume(self) -> tuple():
        if (len(self.command) != 2):
            print("Parameter Missed. Please Input again")
            return (COMMAND_FAILED, 0)
        try:
            volume = float(self.command[1])
            if self.commandChecker.CheckVolume(volume) == COMMAND_SUCC:
                return (COMMAND_SUCC, volume)

            print("Volume Input Error")
            return (COMMAND_FAILED, 0)
        except:
            print("Volume Input Error")
            return (COMMAND_FAILED, 0)


    def __UpdateVolume(self, volume):
        self.newRhythmStatus.volume = volume

    def __UpdateRhythmStatus(self):
        self.newRhythmStatus.isPlaying = 1
        self.newRhythmStatus.isPaused = 0

    def OperateCommand(self) -> tuple():
        if self.rhythmStatus.isPlaying == 0 and self.rhythmStatus.isPaused == 1:
            print("No Rhythm to Play")
            return (COMMAND_FAILED, self.rhythmStatus)

        ret = self.__GetTargetVolume()
        if (ret[0] == COMMAND_FAILED):
            return (COMMAND_FAILED, self.rhythmStatus)
        self.__UpdateVolume(ret[1])
        self._KillOldRhythm()

        self._PlayRhythm(self.newRhythmStatus)
        self.__UpdateRhythmStatus()
        return (COMMAND_SUCC, self.newRhythmStatus)

class ChangeLooptimeAndVolume(RhythmCommander):
    def __GetTargetLooptimeAndVolume(self) -> tuple():
        if (len(self.command) != 3):
            print("Parameter Missed. Please Input again")
            return (COMMAND_FAILED, 0)
        try:
            looptime = int(self.command[1])
            volume = float(self.command[2])
            if (self.commandChecker.CheckVolume(volume) == COMMAND_SUCC and
                self.commandChecker.CheckLooptime(looptime) == COMMAND_SUCC):
                return (COMMAND_SUCC, MusicInfo(None, looptime, volume))

            print("LoopTime Or Volume Input Error")
            return (COMMAND_FAILED, 0)
        except:
            print("LoopTime Or Volume Input Error")
            return (COMMAND_FAILED, 0)

    def __UpdateLoopTimeAndVolume(self, loopTime, volume):
        self.newRhythmStatus.volume = volume

    def __UpdateRhythmStatus(self):
        self.newRhythmStatus.isPlaying = 1
        self.newRhythmStatus.isPaused = 0

    def OperateCommand(self) -> tuple():
        if self.rhythmStatus.isPlaying == 0 and self.rhythmStatus.isPaused == 1:
            print("No Rhythm to Play")
            return (COMMAND_FAILED, self.rhythmStatus)

        ret = self.__GetTargetLooptimeAndVolume()
        if (ret[0] == COMMAND_FAILED):
            return (COMMAND_FAILED, self.rhythmStatus)
        self.__UpdateLoopTimeAndVolume(ret[1].looptime, ret[1].volume)
        self._KillOldRhythm()

        self._PlayRhythm(self.newRhythmStatus)
        self.__UpdateRhythmStatus()
        return (COMMAND_SUCC, self.newRhythmStatus)

class ExitProc(RhythmCommander):
    def OperateCommand(self) -> tuple():
        self._KillOldRhythm()
        time.sleep(0.1) # 为了保证文件删除是在播放进程被kill掉之后 需要有一个较小的延时
        try:
            os.remove(self.pathFile["PreludePath"])
            os.remove(self.pathFile["LoopPath"])
            os.remove(self.pathFile["EpisodePath"])
        except:
            pass

        PID = os.getpid()  # 获取的PID是int
        CommandStream = 'taskkill /pid ' + str(PID) + ' /f'
        os.system(CommandStream)
        return (COMMAND_SUCC, self.rhythmStatus)

class PlayRhythm(RhythmCommander):
    def __SecondSearch(self, nameList):
        print("Input a Aspect to Select a Rhythm From : ")
        No = 1
        for item in nameList:
            print("Rhythm%d : %s" % (No, item))
            No += 1
        newList = []
        newPart = input().lower()
        while newPart == "":
            print("Pleasr Input Again")
            newPart = input().lower()
        for name in nameList:
            if newPart in name:
                newList.append(name)
        return newList

    def __GetRhythmName(self, rhythmNamePart) -> tuple():
        nameList = []
        for key, value in self.timeSheet.items():
            if key.startswith(rhythmNamePart):
                nameList.append(key)
        if (len(nameList) > 1):
            nameList = self.__SecondSearch(nameList)
        if (len(nameList) != 1):
            return (COMMAND_FAILED, "")
        return (COMMAND_SUCC, nameList[0])

    def __GetMusicInfo(self) -> tuple():
        if len(self.command) > 3:
            print("Command Error, Please Input Again")
            return (COMMAND_FAILED, ("", 0, 0))
        rhythmName, looptime, volume = None, None, None
        ret = self.__GetRhythmName(self.command[0])
        if (ret[0] != COMMAND_SUCC):
            return (COMMAND_FAILED, MusicInfo(rhythmName, looptime, volume))
        rhythmName = ret[1]
        try:
            looptime = int(self.command[1])
            if self.commandChecker.CheckLooptime(looptime) != COMMAND_SUCC:
                looptime = DEFAULT_LOOPTIME
        except:
            looptime = DEFAULT_LOOPTIME

        try:
            volume = int(self.command[2])
            if self.commandChecker.CheckVolume(volume) != COMMAND_SUCC:
                volume = DEFAULT_VOLUME
        except:
            volume = DEFAULT_VOLUME
        return (COMMAND_SUCC, MusicInfo(rhythmName, looptime, volume))

    def __UpdateMusicInfo(self, musicInfo):
        self.newRhythmStatus.rhythm = musicInfo.name
        self.newRhythmStatus.loopTime = musicInfo.looptime
        self.newRhythmStatus.volume = musicInfo.volume

    def __UpdateRhythmStatus(self):
        self.newRhythmStatus.isPlaying = 1
        self.newRhythmStatus.isPaused = 0

    def __WriteRhythm(self, rhythmName) -> int:
        print("Rhythm Selected : ",self.timeSheet[rhythmName].name)
        rhythmPath = os.path.join(self.pathFile["RootPath"], self.timeSheet[rhythmName].name)

        Rhythm = AudioSegment.from_mp3(rhythmPath)

        Prelude = Rhythm[: self.timeSheet[rhythmName].loopEnd]
        Loop = Rhythm[self.timeSheet[rhythmName].loopStart: self.timeSheet[rhythmName].loopEnd]
        Epilogue = Rhythm[self.timeSheet[rhythmName].loopStart:]

        try :
            Prelude.export(self.pathFile["PreludePath"], format='mp3')
            Loop.export(self.pathFile["LoopPath"], format='mp3')
            Epilogue.export(self.pathFile["EpisodePath"], format='mp3')
        except Exception as error:
            print("ERROR: Failed to Write Mp3 Files! ")
            return COMMAND_FAILED
        return COMMAND_SUCC

    def OperateCommand(self) -> tuple():
        ret = self.__GetMusicInfo()
        if (ret[0] != COMMAND_SUCC):
            return (COMMAND_FAILED, self.rhythmStatus)

        self.__UpdateMusicInfo(ret[1])

        self._KillOldRhythm()

        if (self.__WriteRhythm(self.newRhythmStatus.rhythm) != COMMAND_SUCC):
            return (COMMAND_FAILED, self.rhythmStatus)
        self._PlayRhythm(self.newRhythmStatus)
        self.__UpdateRhythmStatus()
        return (COMMAND_SUCC, self.newRhythmStatus)

class TestRhythm(RhythmCommander):
    def __SecondSearch(self, nameList):
        print("Input a Aspect to Select a Rhythm From : ")
        No = 1
        for item in nameList:
            print("Rhythm%d : %s" % (No, item))
            No += 1
        newList = []
        newPart = input().lower()
        while newPart == "":
            print("Please Input Again")
            newPart = input().lower()
        for name in nameList:
            if newPart in name:
                newList.append(name)
        return newList

    def __GetRhythmNameAndVolume(self) -> tuple():
        if len(self.command) > 3:
            print("Command Error, Please Input Again")
            return (COMMAND_FAILED, "", 0)
        nameList = []
        for key, value in self.timeSheet.items():
            if key.startswith(self.command[1]):
                nameList.append(key)
        if (len(nameList) > 1):
            nameList = self.__SecondSearch(nameList)
        if (len(nameList) != 1):
            return (COMMAND_FAILED, "")

        try:
            volume = int(self.command[2])
            if self.commandChecker.CheckVolume(volume) != COMMAND_SUCC:
                volume = DEFAULT_VOLUME
        except:
            volume = DEFAULT_VOLUME
        return (COMMAND_SUCC, nameList[0], volume)

    def __ClearRhythmStatus(self):
        self.newRhythmStatus = RhythmStatus()

    def __WriteTestRhythm(self, rhythmName) -> tuple():
        print("Rhythm Selected : ",self.timeSheet[rhythmName].name)
        rhythmPath = os.path.join(self.pathFile["RootPath"], self.timeSheet[rhythmName].name)

        Rhythm = AudioSegment.from_mp3(rhythmPath)

        Loop = Rhythm[self.timeSheet[rhythmName].loopEnd - 5* 1000: self.timeSheet[rhythmName].loopEnd] + \
               Rhythm[self.timeSheet[rhythmName].loopStart: self.timeSheet[rhythmName].loopStart + 5*1000]

        try :
            Loop.export(self.pathFile["LoopPath"], format='mp3')
        except Exception as error:
            print("ERROR: Failed to Write Mp3 Files!" + str(error))
            return COMMAND_FAILED
        return COMMAND_SUCC

    def __TestRhythm(self, volume):
        print("Ready to play")
        subprocess.Popen(["python", "TestRhythm.py", str(2), str(volume), str(10000)], shell=True)
        print("python TestRhythm.py", str(2), str(volume))
        time.sleep(5)

    def OperateCommand(self) -> tuple():
        ret = self.__GetRhythmNameAndVolume()
        if (ret[0] != COMMAND_SUCC):
            return (COMMAND_FAILED, self.rhythmStatus)
        rhythmName, volume = ret[1], ret[2]
        self._KillOldRhythm()

        if (self.__WriteTestRhythm(rhythmName) != COMMAND_SUCC):
            return (COMMAND_FAILED, self.rhythmStatus)
        self.__TestRhythm(volume)
        self.__ClearRhythmStatus()
        return (COMMAND_SUCC, self.newRhythmStatus)

class SheetOperator:
    def __init__(self, pathFile):
        self.Sheet = None
        self.pathFile = pathFile

        self.rhythmStatus = RhythmStatus()

        self.loadSheet()  # Creat Sheet Object

    def loadSheet(self):
        TimeSheet = {}  # A Dic mapping to Music class

        wb = openpyxl.load_workbook(self.pathFile["SheetPath"])
        sheet = wb["Sheet1"]

        for row in sheet[sheet.min_row + 1: sheet.max_row]:
            Rhythm = row[0].value.lower()
            if Rhythm not in TimeSheet:
                NewMusic = Music(row[1].value, row[2].value, row[3].value, row[4].value)
                TimeSheet[Rhythm] = NewMusic
        self.Sheet = TimeSheet
        print("Success to load TimeSheet")

    def __GetCommand(self) -> list:
        return input().lower().strip().split(" ")

    def __GetCommander(self, command) -> RhythmCommander:
        if (command[0] == "help"):
            return ShowHelp(self.Sheet, self.pathFile, self.rhythmStatus, command)
        elif (command[0] == "stop"):
            return StopRhythm(self.Sheet, self.pathFile, self.rhythmStatus, command)
        elif (command[0] == "pause"):
            return PauseRhythm(self.Sheet, self.pathFile, self.rhythmStatus, command)
        elif (command[0] == "play"):
            return RestartRhythm(self.Sheet, self.pathFile, self.rhythmStatus, command)
        elif (command[0] == "find"):
            return ShowAllRhythm(self.Sheet, self.pathFile, self.rhythmStatus, command)
        elif (command[0] == "status"):
            return ShowStatus(self.Sheet, self.pathFile, self.rhythmStatus, command)
        elif (command[0] == "repeat"):
            return ChangeLooptime(self.Sheet, self.pathFile, self.rhythmStatus, command)
        elif (command[0] == "volume"):
            return ChangeVolume(self.Sheet, self.pathFile, self.rhythmStatus, command)
        elif (command[0] == "repeat_volume"):
            return ChangeLooptimeAndVolume(self.Sheet, self.pathFile, self.rhythmStatus, command)
        elif (command[0] == "exit"):
            return ExitProc(self.Sheet, self.pathFile, self.rhythmStatus, command)
        elif (command[0] == "test"):
            return TestRhythm(self.Sheet, self.pathFile, self.rhythmStatus, command)
        else:
            return PlayRhythm(self.Sheet, self.pathFile, self.rhythmStatus, command)

    def __UpdateRhythmStatus(self, newRhythmStatus):
        self.rhythmStatus = newRhythmStatus

    def MainProcess(self):
        while True:
            command = self.__GetCommand()
            if command[0] == "":
                print("Input is None, please input again")
                continue
            commander = self.__GetCommander(command)

            ret = commander.OperateCommand()
            if (ret[0] == COMMAND_SUCC):
                self.__UpdateRhythmStatus(ret[1])


if __name__ == '__main__':
    with open("E:/python/Music/Path.json") as text:
        pathFile = json.load(text)

    pidManager = PidManager()
    pidManager.CleanPidFile()

    rhythmSheet = SheetOperator(pathFile)

    rhythmSheet.MainProcess()


