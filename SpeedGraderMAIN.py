import sys

from PyQt5.QtWidgets import QDialog, QApplication
from PyQt5.QtWidgets import QFileDialog, QMessageBox

from Checkref import speedGrade
from graderUI import Ui_SpeedGraderWindow


class main_window(QDialog):
    def __init__(self):
        super(main_window, self).__init__()
        self.ui = Ui_SpeedGraderWindow()
        self.ui.setupUi(self)
        self.assign_widgets()

        self.lastIndex = None

        # Structure UI widgets into lists for processing use
        # Problem 1 structure
        self.prb1lbl = [self.ui.oneAlbl, self.ui.oneBlbl, self.ui.oneClbl, self.ui.oneDlbl, self.ui.oneElbl,
                        self.ui.oneFlbl,
                        self.ui.oneGlbl, self.ui.oneHlbl]
        self.prb1txt = [self.ui.oneAtxt, self.ui.oneBtxt, self.ui.oneCtxt, self.ui.oneDtxt, self.ui.oneEtxt,
                        self.ui.oneFtxt,
                        self.ui.oneGtxt, self.ui.oneHtxt]
        # Problem 2 structure
        self.prb2lbl = [self.ui.twoAlbl, self.ui.twoBlbl, self.ui.twoClbl, self.ui.twoDlbl, self.ui.twoElbl,
                        self.ui.twoFlbl,
                        self.ui.twoGlbl, self.ui.twoHlbl]
        self.prb2txt = [self.ui.twoAtxt, self.ui.twoBtxt, self.ui.twoCtxt, self.ui.twoDtxt, self.ui.twoEtxt,
                        self.ui.twoFtxt,
                        self.ui.twoGtxt, self.ui.twoHtxt]
        # Problem 3 structure
        self.prb3lbl = [self.ui.threeAlbl, self.ui.threeBlbl, self.ui.threeClbl, self.ui.threeDlbl, self.ui.threeElbl,
                        self.ui.threeFlbl,
                        self.ui.threeGlbl, self.ui.threeHlbl]
        self.prb3txt = [self.ui.threeAtxt, self.ui.threeBtxt, self.ui.threeCtxt, self.ui.threeDtxt, self.ui.threeEtxt,
                        self.ui.threeFtxt,
                        self.ui.threeGtxt, self.ui.threeHtxt]
        # Problem 4 structure
        self.prb4lbl = [self.ui.fourAlbl, self.ui.fourBlbl, self.ui.fourClbl, self.ui.fourDlbl, self.ui.fourElbl,
                        self.ui.fourFlbl,
                        self.ui.fourGlbl, self.ui.fourHlbl]
        self.prb4txt = [self.ui.fourAtxt, self.ui.fourBtxt, self.ui.fourCtxt, self.ui.fourDtxt, self.ui.fourEtxt,
                        self.ui.fourFtxt,
                        self.ui.fourGtxt, self.ui.fourHtxt]
        # Problem 5 structure
        self.prb5lbl = [self.ui.fiveAlbl, self.ui.fiveBlbl, self.ui.fiveClbl, self.ui.fiveDlbl, self.ui.fiveElbl,
                        self.ui.fiveFlbl,
                        self.ui.fiveGlbl, self.ui.fiveHlbl]
        self.prb5txt = [self.ui.fiveAtxt, self.ui.fiveBtxt, self.ui.fiveCtxt, self.ui.fiveDtxt, self.ui.fiveEtxt,
                        self.ui.fiveFtxt,
                        self.ui.fiveGtxt, self.ui.fiveHtxt]
        # Problem 6 structure
        self.prb6lbl = [self.ui.sixAlbl, self.ui.sixBlbl, self.ui.sixClbl, self.ui.sixDlbl, self.ui.sixElbl,
                        self.ui.sixFlbl,
                        self.ui.sixGlbl, self.ui.sixHlbl]
        self.prb6txt = [self.ui.sixAtxt, self.ui.sixBtxt, self.ui.sixCtxt, self.ui.sixDtxt, self.ui.sixEtxt,
                        self.ui.sixFtxt,
                        self.ui.sixGtxt, self.ui.sixHtxt]

        self.groupTxt = [self.prb1txt, self.prb2txt, self.prb3txt, self.prb4txt, self.prb5txt, self.prb6txt]
        self.grouplbl = [self.prb1lbl, self.prb2lbl, self.prb3lbl, self.prb4lbl, self.prb5lbl, self.prb6lbl]

        self.show()

    def assign_widgets(self):
        self.ui.LoadSubmissionsButt.clicked.connect(self.loadSubmissions)
        self.ui.GradeSheetButt.clicked.connect(self.loadGradesheet)
        self.ui.LoadAll.clicked.connect(self.loadProcess)
        self.ui.ExitButt.clicked.connect(self.ExitApp)
        self.ui.listBox.clicked.connect(self.gradePop)
        self.ui.ScrollUp.clicked.connect(self.clickUp)
        self.ui.ScrollDown.clicked.connect(self.clickDown)
        self.ui.AbsentWorkCheck.clicked.connect(self.quickCheck)
        self.ui.SaveExcel.clicked.connect(self.saveAndQuit)
        self.ui.LoadPDF.clicked.connect(self.getPDF)

    # Initializes the UI for the gradesheet layout and student names
    def loadProcess(self):
        self.massDisable()
        self.ui.listBox.clear()
        try:
            self.myGrader = speedGrade(self.ui.GradeSheetInput.text(), self.ui.HomeworkSubmissions.text())
            self.ui.listBox.addItems(self.myGrader.names)
            self.ui.listBox.setCurrentRow(0)
            self.gradePop()
            # Populate the problem headers into the UI
            hdrCount = 0
            for prblms in range(len(self.myGrader.problemChecks[0])):
                for hdrs in range(1 + self.myGrader.problemChecks[1][prblms] - self.myGrader.problemChecks[0][prblms]):
                    self.grouplbl[prblms][hdrs].setText(self.myGrader.probHead[hdrCount])
                    self.groupTxt[prblms][hdrs].setEnabled(True)
                    hdrCount += 1
        except:
            self.bad_file()

    # Populates active cells with student's grades
    def gradePop(self):
        self.recommit()
        scores = self.myGrader.pullScores(self.ui.listBox.currentRow())
        hdrCount = 0
        for prblms in range(len(self.myGrader.problemChecks[0])):
            for hdrs in range(1 + self.myGrader.problemChecks[1][prblms] - self.myGrader.problemChecks[0][prblms]):
                if scores[hdrCount] == None:
                    self.groupTxt[prblms][hdrs].setText('')
                else:
                    self.groupTxt[prblms][hdrs].setText("{}".format(scores[hdrCount]))
                # Saves the index associated with the last set of populated grades
                hdrCount += 1
        self.lastIndex = self.ui.listBox.currentRow()

    # Recommits input grades to the spreadsheet object
    def recommit(self):
        if self.lastIndex == None:
            pass
        else:
            for prblms in range(len(self.myGrader.problemChecks[0])):
                for hdrs in range(self.myGrader.problemChecks[0][prblms], 1 + self.myGrader.problemChecks[1][prblms]):
                    self.myGrader.mySheet.cell(self.lastIndex + 2, hdrs).value = self.groupTxt[prblms][
                        hdrs - self.myGrader.problemChecks[0][prblms]].text()

    # Disables all input boxes in the problems and reverts all label boxes to -
    def massDisable(self):
        for group in self.groupTxt:
            for txtBx in group:
                txtBx.setEnabled(False)
        for group in self.grouplbl:
            for lblBx in group:
                lblBx.setText('-')

    # Requests PDF from the speed grader object for viewing
    def getPDF(self):
        if self.ui.listBox.currentRow() < 0:
            pass
        else:
            returnBool = self.myGrader.pullPDF(self.ui.listBox.currentRow())
            if returnBool == False:
                msg = QMessageBox()
                msg.setIcon(QMessageBox.Information)
                msg.setText('PDF Not Found')
                msg.exec()

    # Moves index up when up arrow key selected
    def clickUp(self):
        if self.ui.listBox.currentRow() <= 0:
            pass
        else:
            self.ui.listBox.setCurrentRow(self.ui.listBox.currentRow() - 1)
            self.gradePop()

    # Moves index down when down arrow key selected
    def clickDown(self):
        if self.ui.listBox.currentRow() < 0:
            pass
        elif self.ui.listBox.currentRow() == (len(self.myGrader.names) - 1):
            pass
        else:
            self.ui.listBox.setCurrentRow(self.ui.listBox.currentRow() + 1)
            self.gradePop()

    # Calls the zeroize function to find and mark unsubmitted pdfs as zeroes
    def quickCheck(self):
        if self.ui.listBox.currentRow() != -1:
            self.myGrader.zeroize()
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Information)
            msg.setText('{} Missing Submissions Found'.format(self.myGrader.absCount))
            msg.exec()

    # Notifies the user that the requested file does not exist or the directory is bad
    def bad_file(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Warning)
        msg.setText('Unable to process the selected file')
        msg.exec()

    # Opens the directory input for the submissions folder
    def loadSubmissions(self):
        self.ui.HomeworkSubmissions.setText(QFileDialog.getExistingDirectory())

    # Opens the directory input for the gradesheet
    def loadGradesheet(self):
        self.ui.GradeSheetInput.setText(QFileDialog.getOpenFileName(filter='*.xlsx')[0])

    # Calls the speed grader to save all locally stored excel data
    def saveAndQuit(self):
        if self.ui.listBox.currentRow() != -1:
            self.myGrader.save()
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Information)
            msg.setText('Spreadsheet Saved')
            msg.exec()

    # Exits out of the entire UI
    def ExitApp(self):
        app.exit()


if __name__ == "__main__":
    app = QApplication.instance()
    if not app:
        app = QApplication(sys.argv)
    app.aboutToQuit.connect(app.deleteLater)
    main_win = main_window()
    sys.exit(app.exec_())