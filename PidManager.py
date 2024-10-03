
class PidManager:
    def __init__(self):
        self.pidFilePath = "E:/python/Music/PID.txt"

    def CleanPidFile(self):
        file = open(self.pidFilePath, "w")
        file.close()

    def GetPid(self) -> str:
        file = open(self.pidFilePath, "r")
        pid = file.readline()
        file.close()
        return pid

    def RecordPid(self, pid):
        file = open(self.pidFilePath, "w")  # 获取PID并输出
        file.write(pid)
        file.close()