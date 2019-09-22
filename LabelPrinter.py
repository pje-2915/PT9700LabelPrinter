import subprocess
import xxsubtype
import shlex
from subprocess import Popen, PIPE

class LabelPrinter:

    def __init__(self, speed=0):
        self.cmdStart = bytearray.fromhex('1B6961001B40')  # use Esc P, initialise
        self.cmdLandscape = bytearray.fromhex('1B694C00')  # rotated printing off
        self.cmdFormat = bytearray.fromhex('1B696D0700')  # minimum margin
        self.cmdTextOps = bytearray.fromhex('1B6B001B5834')  # Helsinki, 04 -> 9 pt
        self.cmdCutSetting = bytearray.fromhex('1B694306')  # chain print + 1/2 cut
        self.cmdLF = bytearray.fromhex('0A')
        self.cmdPrint = bytearray.fromhex('0C')
        self.botanicalName = ""
        self.collectionNumber = ""
        self.localityData = ""
        self.ballParkFontRatio = 2  # ratio of botanical name font to locality font
        self.name = ""

    def sendPrintSequence(self):
        splitname = self.lineSplitter(self.name)

        monsterprintbuffer = self.cmdStart + self.cmdLandscape + self.cmdFormat+self.cmdTextOps \
        +bytearray(splitname,'latin-1') + self.cmdLF + bytearray(self.localityData,'latin-1') \
        + self.cmdCutSetting + self.cmdPrint

        args = shlex.split("/usr/bin/lpr -l -P Brother_PT_9700PC")
        p = Popen(args, stdin=PIPE)
        p.stdin.write(monsterprintbuffer)
        p.communicate()

    def lineSplitter(self, longLine: str):

        # Roughly divide the text in two at a convenient word boundary
        total_len = len(longLine)
        words = longLine.split(' ')
        word_index = 0
        len_so_far = 0
        tmp_length = 0

        for word in words:
            tmp_length = len(word) + 1  # word length + space (this is rough!)
            if len_so_far + tmp_length > total_len/2:
                if len_so_far + tmp_length/2 < total_len / 2:
                    word_index += 1
                break
            len_so_far += tmp_length
            word_index += 1

        print(word_index)
        split_index = word_index

        word_index = 0
        outstring = ""
        for word in words:
            print(word)
            if word_index == split_index and word_index > 0:
                outstring += chr(0x0A)
            else:
                outstring += " "
            outstring += word
            word_index += 1
        print("Out string:",outstring, split_index)
        return outstring

if __name__ == '__main__':

    myLP = LabelPrinter()
    while True:
        myLP.botanicalName = input("Botanical Name: ")
        myLP.collectionNumber = input("Collection Number: ")
        myLP.localityData = input("Locality Data: ")
        myLP.name = myLP.botanicalName+" "+myLP.collectionNumber

        print(myLP.name, myLP.localityData)
        myLP.sendPrintSequence()
        # myLP.lineSplitter(myLP.botanicalName)