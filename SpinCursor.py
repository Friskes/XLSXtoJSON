from threading import Thread
from sys import stdout
from os import name as os_name
from time import sleep
from unicodedata import lookup

class SpinCursor(Thread):
    """ A console spin cursor class """

    def __init__(self,
                 msg='',
                 del_msg_after_stop=False,
                 maxspin=0,
                 minspin=10,
                 speed=2,
                 animType='sticks'):

        # Count of a spin
        self.count = 0
        self.out = stdout
        self.flag = False
        self.max = maxspin
        self.min = minspin
        # Any message to print first ?
        self.msg = msg
        # Complete printed string
        self.string = ''
        # Speed is given as number of spins a second
        # Use it to calculate spin wait time
        self.waittime = 1.0 / float(speed * 4)
        self.animType = animType
        self.del_msg = del_msg_after_stop

        if os_name == 'posix':
            if self.animType == 'sticks':
                self.spinchars = (lookup('FIGURE DASH'), u'\\ ', u'| ', u'/ ')
            elif self.animType == 'dots':
                self.spinchars = (u'   ', u'.  ', u'.. ', u'...')
            elif self.animType == 'nums':
                self.spinchars = (u'0', u'1', u'2', u'3')
        else:
            # The unicode dash character does not show
            # up properly in Windows console.
            if self.animType == 'sticks':
                self.spinchars = (u'—', u'\\ ', u'| ', u'/ ')
            elif self.animType == 'dots':
                self.spinchars = (u'   ', u'.  ', u'.. ', u'...')
            elif self.animType == 'nums':
                self.spinchars = (u'0', u'1', u'2', u'3')

        Thread.__init__(self, name="SpinCursor Thread")

    def spin(self):
        """ Perform a single spin """

        for x in self.spinchars:

            if self.animType == 'sticks' or self.animType == 'nums':
                if self.msg == '':
                    self.string = self.msg + x + "\r"
                else:
                    self.string = self.msg + ' ' + x + "\r"

            elif self.animType == 'dots':
                self.string = self.msg + x + "\r"

            if os_name == 'posix':
                self.out.write(self.string.encode('utf-8'))
            else:
                self.out.write(self.string)

            self.out.flush()
            sleep(self.waittime)

    def run(self):
        while (not self.flag) and ((self.count < self.min) or (self.count < self.max)):
            self.spin()
            self.count += 1

        # Clean up display...
        if self.del_msg:
            self.out.write(" " * len(self.string) + "\r")
        else:
            if self.msg == '':
                self.out.write(self.msg + "✔  \n")
            else:
                self.out.write(self.msg + " ✔ \n")

    def stop(self):
        self.flag = True
