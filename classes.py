class App :

    def __init__(self, code, name, ac, phase):
        self.code = code
        self.name = name
        self.appCustodian = ac
        self.phase = phase
        self.slaStart = ""
        self.slaEnd = ""
        self.frequency = ""
        self.status = ""

    def add_slaInfo(self, slaStart, slaEnd, frequency) :
        self.slaStart = slaStart
        self.slaEnd = slaEnd
        self.frequency = frequency

    def set_status(self, st) :
        self.status = st

class Contact :

    def __init__(self, app, initContact, lastContact, replyDate, target) :
        self.app = app
        self.first = initContact
        self.last = lastContact
        self.reply = replyDate
        self.target = target
        self.status = ""
        self.comment = ""

    def set_comment(self, cmt) :
        self.comment = cmt

    def set_status(self, st) :
        self.status = st
