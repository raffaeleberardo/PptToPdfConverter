import comtypes.client
import os
import glob

def PPTtoPDF(inputFileName, outputFileName, formatType = 32):
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1
    if outputFileName[-3:] != 'pdf':
        outputFileName = outputFileName + '.pdf'
    deck = powerpoint.Presentations.Open(inputFileName)
    deck.SaveAs(outputFileName, formatType) #formatType = 32 for ppt to pdf
    deck.Close()
    powerpoint.Quit()

dir = None
while dir is None or not os.path.isdir(dir):
    try:
        dir = input('Insert absolute path where there are all your ppt file you want to convert:\n')
        if os.path.isdir(dir):
            os.chdir(dir)
            print(dir)
            for file in glob.glob("*.ppt"):
                print(file)
                print(file[:-4])
                PPTtoPDF(dir + "\\" + file, dir + "\\" + file[:-4])
    except COMError as ce:
        target_error = ce.args  # this is a tuple
        if target_error[1] == 'Call was rejected by callee.':
            self.acad.doc.SendCommand("Chr(3)")


