from cgi import test
from pystray import Icon, Menu, MenuItem
from PIL import Image

class SystemTray:
	#Initialises the system tray
	def __init__(self):
		#system tray image loading
		image = Image.open('py.ico')
		image.load()
		
		#Makes a menu for the system tray icon
		menu = Menu(
			MenuItem('Start Recognition', self.onStartRecognition),
			MenuItem('Stop Recognition', self.onStopRecognition),
			MenuItem('Exit', self.exitProgram))
		#Creates the icon and makes it visible
		self.icon = Icon(name='Speech Transcriber', title='Speech Transcriber', icon=image, menu=menu)
		self.icon.visible = True
	
	#Starts the system tray
	def startSysTray(self):		
		self.icon.run()
	
	#Starts recognition when the start button is clicked on the menu
	def onStartRecognition(self, icon):
		print("start")
	
	#Stops recognition when the stop button is clicked on the menu
	def onStopRecognition(self, icon):
		print("stop")

	#Closes the program when the quit button is clicked in the menu
	def exitProgram(self, icon):
        # print("Exiting program...")
		self.icon.stop()