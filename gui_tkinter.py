import Tkinter, Tkconstants, tkFileDialog
RED = "#f44"
GREEN = "#4b4"
global INITIAL_DIRECTORY
INITIAL_DIRECTORY = "/home/akshay/Public/AutomaticDocumentEditor/AXA"
import monitor
from utilities import check_expired

class TkFileDialogExample(Tkinter.Tk):

    frame_header = None
    frame_footer = None
    label_directoryname = None
    button_askdirectory = None
    button_proceed = None
    button_close = None
    frame_footer = None
    label_status = None
    label_dont_quit = None
    button_startservice = None
    label_expiredtext = None

    def unpack_all(self):
        vars = (
            self.frame_header,
            self.frame_footer,
            self.label_directoryname,
            self.button_askdirectory,
            self.button_proceed,
            self.button_close,
            self.frame_footer,
            self.label_status,
            self.label_dont_quit,
            self.button_startservice,
            self.label_expiredtext
        )
        for var in vars:
            if var != None:
                var.pack_forget()

    def __init__(self):
        Tkinter.Tk.__init__(self)

        self.frame_header = Tkinter.Frame(self, width=256, height=10)
        self.frame_header.pack(fill=None, expand=False)
        self.frame_footer = Tkinter.Frame(self, width=256, height=10)

        # TODO Change for windows
        self.directoryname = INITIAL_DIRECTORY
        # options for UI
        self.button_opt = {'fill': Tkconstants.BOTH, 'padx': 5, 'pady': 5}

        self.label_directoryname = Tkinter.Label(self, text='No Folder Selected')
        self.label_directoryname.pack(**self.button_opt)

        self.label_expiredtext = Tkinter.Label(self, text='Product Expired')

        self.button_askdirectory = Tkinter.Button(self, text='Select Folder', command=self.askdirectory)
        self.button_askdirectory.pack(**self.button_opt)

        self.button_proceed = Tkinter.Button(self, text='Proceed', command=self.proceed)
        self.button_proceed.pack(**self.button_opt)

        self.button_close = Tkinter.Button(self, text = 'Close', command=self.closeservice)
        self.button_close.pack()

        self.frame_footer.pack()

        self.label_status = Tkinter.Label(self, text = 'Service Status: Not Running', fg = RED)
        self.label_dont_quit = Tkinter.Label(self, text = 'Please don\'t close this window')
        self.button_startservice = Tkinter.Button(self, text = 'Start Service', command=self.start_service)

        # defining options for opening a directory
        self.dir_opt = options = {}
        options['initialdir'] = INITIAL_DIRECTORY
        options['mustexist'] = True
        options['parent'] = self
        options['title'] = 'This is a title'

        if check_expired():
            self.expired_view()
            return



    def askdirectory(self):
        if check_expired():
            self.expired_view()
            return
        self.directoryname = tkFileDialog.askdirectory(**self.dir_opt)
        self.label_directoryname['text'] = 'Folder path: ' + self.directoryname
        self.button_askdirectory['text'] = 'Change Folder'

    def proceed(self):
        if check_expired():
            self.expired_view()
            return
        self.unpack_all()
        print 'Selected Folder:', self.directoryname
        self.frame_header.pack()
        self.label_status.pack()
        self.button_startservice.pack()
        self.button_close.pack()
        self.frame_footer.pack()

    def start_service(self):
        if check_expired():
            self.expired_view()
            return

        print "**************************************"
        print self.directoryname.strip()
        monitor.main(self.directoryname.strip(), self)
        self.label_status['text'] = 'Service Status: Running'
        self.label_status['fg'] = GREEN
        self.button_close['text'] = 'Stop Service'

        self.unpack_all()

        self.frame_header.pack()
        self.label_dont_quit.pack()
        self.button_close.pack()
        self.frame_footer.pack()

    def expired_view(self):
        self.unpack_all()
        self.frame_header.pack(fill=None, expand=False)
        self.label_expiredtext.pack(**self.button_opt)
        self.button_close['text'] = 'Close'
        self.button_close.pack()
        self.frame_footer.pack(fill=None, expand=False)

    def closeservice(self):
        self.destroy()


if __name__=='__main__':
    root = TkFileDialogExample()
    root.wm_title('Automatic Document Editor')
    root.mainloop()
