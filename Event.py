class Event(object):
    startTime = ""
    endTime = ""
    opt_Cherry = []
    opt_URL = []
    opt_Keywords = []
    limit = 70

    def setStartTime(self, value):
        self.startTime = value
    def getStartTime(self):
        return self.startTime

    def setEndTime(self, value):
        self.endTime = value
    def getEndTime(self):
        return self.endTime

    def setOpt_Cherry(self, value):
        self.opt_Cherry.extend(value)
    def getOpt_Cherry(self):
        return self.opt_Cherry

    def setOpt_URL(self, value):
        self.opt_URL.extend(value)
    def getOpt_URL(self):
        return self.opt_URL

    def setOpt_Keywords(self, value):
        self.opt_Keywords.extend(value)
    def getOpt_Keywords(self):
        return self.opt_Keywords