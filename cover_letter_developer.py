import sys
from PySide2.QtWidgets import QApplication, QDialog, QLineEdit, QPushButton,QVBoxLayout, QLabel, QWidget
from docx import Document
from docx.shared import Inches

class Form(QDialog):

    def __init__(self, parent=None):
        super(Form, self).__init__(parent)
        #set the size
        #Creat widgets
        self.setWindowTitle("Cover Letter Developer")
        self.label1 = QLabel('Input Company Name')
        self.edit1 = QLineEdit("")
        self.label2 = QLabel('Input Position Title')
        self.edit2 = QLineEdit("")
        self.label3 = QLabel('How did you get introduced to the company?')
        self.edit3 = QLineEdit("")
        self.label4 = QLabel('What skills do you have that would help the COOP/Internship')
        self.edit4 = QLineEdit("")
        self.button = QPushButton("Develop")
        # Creat layout and add widgets
        layout  = QVBoxLayout()
        layout.addWidget(self.label1)
        layout.addWidget(self.edit1)
        layout.addWidget(self.label2)
        layout.addWidget(self.edit2)
        layout.addWidget(self.label3)
        layout.addWidget(self.edit3)
        layout.addWidget(self.label4)
        layout.addWidget(self.edit4)
        layout.addWidget(self.button)
        #set dialog layout
        self.setLayout(layout)
        self.button.clicked.connect(self.coverlet)


    def coverlet(self):
        name = self.edit1.text()
        pos = self.edit2.text()
        intro = self.edit3.text()
        skills = self.edit4.text()
        mytext = """
        Dear """ + name + """’s Hiring Team,
        \n
            """ + """   """ + """   I am writing to apply to the """ + pos + """ Intern/COOP position at """ + name + """. I am a 4th year at Wentworth Institute of Technology, pursuing a Bachelor of Science degree in Electro-mechanical Engineering. The Electro-mechanical Engineering program combines the technical disciplines of Electrical and Mechanical Engineering. """ + intro + """ 
            
            """+ """As an intern at """ + name + """ , I will bring my toolset of """ + skills + """. Additionally I have experience in quality and reliability of electronic circuit systems through the tests that I have done when I was Analog Devices like shock, high voltage, HALT testing. Along with developing reliability testers that I programmed using LabView(a graphical programming language). My C programming and Python experience is from a project that I have done for a Junior Design Project and you can see the pictures through my personal website list below.

            """ + """   """ + """   As an engineering student, the most valuable thing that I have currently learned about myself is that when faced with a difficult problem I may initially fail, but I don’t quit until I eventually solve the problem. I am a quick learner and will be a good asset to """ + name + """. Wentworth Institute of Technology incorporates COOPS/internships as part of its curriculum, and, therefore, I would be available to work full time throughout the summer for a minimum of 14 weeks. I would be honored to intern for """ + name + """ and gain experience in engineering and further """+ name +""" initiative. has a reputation for excellence, and I value your commitment to making the world a better and safer place.

            """ + """   """ + """   You may contact me by phone, email or my personal website, which I have supplied below. Thank you for your time and consideration.

        """

        anothertext = """ 
Respectfully yours,
Martynas Baranauskas
baranauskasm@wit.edu
781-572-9775
Personal Website: https://baranauskasm.wixsite.com/mysite
or scan QR code with smartphone camera
        """

        document = Document()
        p = document.add_paragraph(mytext)
        g = document.add_paragraph(anothertext)
        k = document.add_picture('qr_code.png', width=Inches(0.7))
        # document.add_page_break()

        # the saving of the document and the path to the
        filename = name + '_' + pos + '_baranauskas_.docx'
        # filepath = r'C:\Users\baranauskasm\Desktop\COOP Stuff\Summer 2020 COOP (future)\cover letters\automated cover letters'
        document.save(filename)
        print("-----------------------------------------------------")
        print(name + "_" + pos + "_baranauskas.doxc document was developed")
        print("------------------------------------------------------")

        #clear the form for another submition
        self.edit1.clear()
        self.edit2.clear()
        self.edit3.clear()
        self.edit4.clear()

if __name__ == '__main__':
    #or you can do a automatic one with something like
    # Create the Qt Application
    app = QApplication(sys.argv)
    # Create and show the form
    form = Form()
    #the size of the gui
    form.resize(1300,250)
    form.show()
    # Run the main Qt loop
    sys.exit(app.exec_())
