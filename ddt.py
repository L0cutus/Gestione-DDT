#!/usr/bin/env python
# -*- coding: utf-8 -*-

from __future__ import division
from __future__ import print_function
from __future__ import unicode_literals
from future_builtins import *

import os
import sys
import subprocess

from PyQt4.QtCore import PYQT_VERSION_STR, QDate, QFile
from PyQt4.QtCore import QRegExp, QString, QVariant, Qt
from PyQt4.QtCore import SIGNAL, QModelIndex, QSettings
from PyQt4.QtCore import QSize, QPoint

from PyQt4.QtGui  import QApplication, QCursor, QDateEdit
from PyQt4.QtGui  import QDialog, QMainWindow, QHBoxLayout
from PyQt4.QtGui  import QLabel, QLineEdit, QMessageBox, QPixmap
from PyQt4.QtGui  import QTabWidget, QPushButton, QRegExpValidator
from PyQt4.QtGui  import QStyleOptionViewItem, QTableView, QVBoxLayout
from PyQt4.QtGui  import QDataWidgetMapper, QTextDocument, QStyle
from PyQt4.QtGui  import QColor, QBrush, QTextOption
from PyQt4.QtGui  import QItemSelectionModel,QStandardItemModel
from PyQt4.QtGui  import QAbstractItemView, QIntValidator
from PyQt4.QtGui  import QDoubleValidator, QIcon, QFileDialog, QItemDelegate

from PyQt4.QtSql  import QSqlDatabase, QSqlQuery, QSqlRelation
from PyQt4.QtSql  import QSqlRelationalDelegate, QSqlRelationalTableModel
from PyQt4.QtSql  import QSqlTableModel

import ddt_ui
import aboutddt

# Definizione 6 righe intestazione ditta
r1 = "TIME di Stefano Zamprogno"                        # ragione sociale
r2 = "Via A.Bonetto, 6 31044 Montebelluna (TV)"         # es. indirizzo
r3 = "Cod.Fisc. ZMPSFN66T26F443D"                       # es. Codice Fiscale
r4 = "P.IVA: 02297230266 Reg. Impr. TV n'131843"        # es. Partita Iva
r5 = "Tel. 04231900335"                                 # es. telefono

# Definizione degli 'id' usati poi come colonne nelle tabelle ecc...
MID, MDATA, MDDT, MIDCLI, MCAU, MNOTE = range(6)
SID, SQT, SDESC, SMID = range(4)
CID, CRAGSOC, CIND, CPIVA, CCF, CTEL, CFAX, CCELL, CEMAIL = range(9)

DATEFORMAT = "dd/MM/yyyy"

# usate per il salvataggio dei settings dell'applicazione
DDTORG = "TIME di Stefano Z."
DDTAPP = "Gestione DDT"
DDTDOMAIN = "zamprogno.it"

class MyQLineEdit(QLineEdit):
    def __init__(self, parent=None):
        super(MyQLineEdit, self).__init__(parent)

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Down:
            self.emit(SIGNAL("keyDownPressEvent()"))
            return
        return QLineEdit.keyPressEvent(self, event)

class MyQSqlRelationalDelegate(QSqlRelationalDelegate):
    def __init__(self, parent=None):
        super(MyQSqlRelationalDelegate, self).__init__(parent)

    def createEditor(self, parent, option, index):
        if index.column() == SQT:
            editor = MyQLineEdit(parent)
            validator = QIntValidator(self)
            editor.setValidator(validator)
            editor.setAlignment(Qt.AlignRight|Qt.AlignVCenter)
            self.connect(editor, SIGNAL("keyDownPressEvent()"),
                        self.gestEvt)
            return editor
        elif index.column() == SDESC:
            editor = MyQLineEdit(parent)
            editor.setAlignment(Qt.AlignRight|Qt.AlignVCenter)
            self.connect(editor, SIGNAL("keyDownPressEvent()"),
                        self.gestEvt)
            return editor

        return QSqlRelationalDelegate.createEditor(self, parent,
                                                    option, index)

    def gestEvt(self):
        editor = self.sender()
        if isinstance(editor, (MyQLineEdit)):
            self.emit(SIGNAL("commitData(QWidget*)"), editor)
            self.emit(SIGNAL("addDettRecord()"))

class MainWindow(QMainWindow, ddt_ui.Ui_MainWindow):

    FIRST, PREV, NEXT, LAST = range(4)

    def __init__(self, parent=None):
        super(MainWindow, self).__init__(parent)

        self.setupUi(self)
        self.setupMenu()
        self.restoreWinSettings()

        self.editindex = None
        self.filename = None
        self.db = QSqlDatabase.addDatabase("QSQLITE")

        self.loadInitialFile()
        self.setupUiSignals()

    def mmUpdate(self):
        row = self.mapper.currentIndex()
        id = self.mModel.data(self.mModel.index(row,MID)).toString()
        self.sModel.setFilter("mmid=%s" % id)
        self.sModel.select()
        self.sTableView.setColumnHidden(SID, True)
        self.sTableView.setColumnHidden(SMID, True)


    def addDdtRecord(self):
        if not self.db.isOpen():
            self.statusbar.showMessage(
                "Database non aperto...",
                5000)
            return
        row = self.mModel.rowCount()
        self.mapper.submit()
        self.mModel.insertRow(row)
        self.mapper.setCurrentIndex(row)
        self.dateEdit.setDate(QDate.currentDate())
        self.dateEdit.setFocus()
        self.mmUpdate()

    def delDdtRecord(self):
        if not self.db.isOpen():
            self.statusbar.showMessage(
                "Database non aperto...",
                5000)
            return
        row = self.mapper.currentIndex()
        if row == -1:
            self.statusbar.showMessage(
                        "Nulla da cancellare...",
                        5000)
            return
        record = self.mModel.record(row)
        id = record.value(MID).toInt()[0]
        ddt = record.value(MDDT).toString()
        if(QMessageBox.question(self, "Cancella Scaffale",
                    "Vuoi cancellare il ddt N': {0} ?".format(ddt),
                    QMessageBox.Yes|QMessageBox.No) ==
                    QMessageBox.No):
            self.statusbar.showMessage(
                        "Cancellazione ddt annullata...",
                        5000)
            return
        # cancella scaffale
        self.mModel.removeRow(row)
        self.mModel.submitAll()
        if row + 1 >= self.mModel.rowCount():
            row = self.mModel.rowCount() - 1
        self.mapper.setCurrentIndex(row)
        if self.mModel.rowCount() == 0:
            self.cauLineEdit.setText(QString(""))
            self.noteLineEdit.setText(QString(""))
            self.ddtLineEdit.setText(QString(""))
            self.cliComboBox.setCurrentIndex(-1)

        # cancella tutti gli articoli che si riferiscono
        # allo scaffale cancellato
        self.sModel.setFilter("mmid=%s" % id)
        self.sModel.select()
        self.sModel.removeRows(0, self.sModel.rowCount())
        self.sModel.submitAll()
        self.statusbar.showMessage(
                        "Cancellazione eseguita...",
                        5000)
        self.mmUpdate()

    def addDettRecord(self):
        if not self.db.isOpen():
            self.statusbar.showMessage(
                "Database non aperto...",
                5000)
            return
        rowddt = self.mapper.currentIndex()
        record = self.mModel.record(rowddt)
        masterid = record.value(MID).toInt()[0]
        if masterid < 1:
            self.statusbar.showMessage(
                "Scaffale non valido o non confermato...",
                5000)
            self.dateEdit.setFocus()
            return
        # aggiunge la nuova riga alla vista
        self.sModel.submitAll()
        self.sModel.select()
        row = self.sModel.rowCount()
        self.sModel.insertRow(row)
        self.sModel.setData(self.sModel.index(row, SMID),
                                                QVariant(masterid))
        self.sModel.setData(self.sModel.index(row, SQT),
                                                QVariant(1))
        self.sModel.setData(self.sModel.index(row, SDESC),
                                                QVariant(""))
        self.editindex = self.sModel.index(row, SQT)
        self.sTableView.setCurrentIndex(self.editindex)
        self.sTableView.edit(self.editindex)

    def delDettRecord(self):
        if not self.db.isOpen():
            self.statusbar.showMessage(
                "Database non aperto...",
                5000)
            return
        selrows = self.sItmSelModel.selectedRows()
        if not selrows:
            self.statusbar.showMessage(
                "No articles selected to delete...",
                5000)
            return
        if(QMessageBox.question(self, "Cancellazione righe",
                "Vuoi cancellare: {0} righe?".format(len(selrows)),
                QMessageBox.Yes|QMessageBox.No) ==
                QMessageBox.No):
            return
        QSqlDatabase.database().transaction()
        query = QSqlQuery()
        query.prepare("DELETE FROM ddtslave WHERE id = :val")
        for i in selrows:
            if i.isValid():
                query.bindValue(":val", QVariant(i.data().toInt()[0]))
                query.exec_()
        QSqlDatabase.database().commit()
        self.sModel.revertAll()
        self.mmUpdate()


    def setupMenu(self):
        # AboutBox
        self.connect(self.action_About, SIGNAL("triggered()"),
                    self.showAboutBox)
        # FileNew
        self.connect(self.action_New_File, SIGNAL("triggered()"),
                    self.newFile)

        # FileLoad
        self.connect(self.action_Load_File, SIGNAL("triggered()"),
                    self.openFile)

        # Edit Customers
        self.connect(self.action_Add_Customers, SIGNAL("triggered()"),
                    self.editCustomers)

    def editCustomers(self):
        relpath = os.path.dirname(__file__)
        if relpath:
            relpath = "%s/" % relpath
        subprocess.call(['python',os.path.join("%s../clienti/" %
                                        relpath, "clienti.py")])
        self.setupModels()
        self.setupMappers()
        self.setupTables()
        self.mmUpdate()

    def showAboutBox(self):
        dlg = aboutddt.AboutBox(self)
        dlg.exec_()

    def creaStrutturaDB(self):
        query = QSqlQuery()
        if not ("ddtmaster" in self.db.tables()):
            if not query.exec_("""CREATE TABLE ddtmaster (
                                id INTEGER PRIMARY KEY AUTOINCREMENT UNIQUE NOT NULL,
                                data DATE NOT NULL,
                                ddt VARCHAR(50) NOT NULL,
                                idcli INTEGER NOT NULL,
                                causale VARCHAR(200),
                                note VARCHAR(200))"""):
                QMessageBox.warning(self, "Gestione DDT",
                                QString("Creazione tabella master fallita!"))
                return False

        if not ("ddtslave" in self.db.tables()):
            if not query.exec_("""CREATE TABLE ddtslave (
                                id INTEGER PRIMARY KEY AUTOINCREMENT UNIQUE NOT NULL,
                                qt INTEGER NOT NULL DEFAULT '1',
                                desc VARCHAR(200) NOT NULL,
                                mmid INTEGER NOT NULL,
                                FOREIGN KEY (mmid) REFERENCES master)"""):
                QMessageBox.warning(self, "Gestione DDT",
                                QString("Creazione tabella slave fallita!"))
                return False
            QMessageBox.information(self, "Gestione DDT",
                                QString("Database Creato!"))
            return True


        if not ("clienti" in self.db.tables()):
            if not query.exec_("""CREATE TABLE clienti (
                                id INTEGER PRIMARY KEY AUTOINCREMENT UNIQUE NOT NULL,
                                ragsoc VARCHAR(200) NOT NULL,
                                indirizzo VARCHAR(200) NOT NULL,
                                piva VARCHAR(15),
                                cf VARCHAR(15),
                                tel VARCHAR(30),
                                fax VARCHAR(30),
                                cell VARCHAR(30),
                                email VARCHAR(50))"""):
                QMessageBox.warning(self, "Gestione DDT",
                                QString("Creazione tabella clienti fallita!"))
                return False
        return True

    def loadFile(self, fname=None):
        if fname is None:
            return
        if self.db.isOpen():
            self.db.close()
        self.db.setDatabaseName(QString(fname))
        if not self.db.open():
            QMessageBox.warning(self, "Gestione DDT",
                                QString("Database Error: %1")
                                .arg(db.lastError().text()))
        else:
            if not self.creaStrutturaDB():
                return
            self.filename = unicode(fname)
            self.setWindowTitle("Gestione DDT - %s" % self.filename)
            self.setupModels()
            self.setupMappers()
            self.setupTables()
            #self.setupItmSignals()
            self.restoreTablesSettings()
            self.mmUpdate()


    def loadInitialFile(self):
        settings = QSettings()
        fname = unicode(settings.value("Settings/lastFile").toString())
        if fname and QFile.exists(fname):
            self.loadFile(fname)


    def openFile(self):
        dir = os.path.dirname(self.filename) \
                if self.filename is not None else "."
        fname = QFileDialog.getOpenFileName(self,
                    "Gestione DDT - Scegli database",
                    dir, "*.db")
        if fname:
            self.loadFile(fname)


    def newFile(self):
        dir = os.path.dirname(self.filename) \
                if self.filename is not None else "."
        fname = QFileDialog.getSaveFileName(self,
                    "Gestione DDT - Scegli database",
                    dir, "*.db")
        if fname:
            self.loadFile(fname)

    def restoreWinSettings(self):
        settings = QSettings()
        self.restoreGeometry(
                settings.value("MainWindow/Geometry").toByteArray())

    def restoreTablesSettings(self):
        settings = QSettings(self)
        # per la tablelview
        for column in range(1, self.sModel.columnCount()-1):
            width = settings.value("Settings/sTableView/%s" % column,
                                    QVariant(60)).toInt()[0]
            self.sTableView.setColumnWidth(column,
                                        width if width > 0 else 60)

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Down:
            self.addDettRecord()
        else:
            QMainWindow.keyPressEvent(self, event)

    def closeEvent(self, event):
        self.mapper.submit()
        settings = QSettings()
        settings.setValue("MainWindow/Geometry", QVariant(
                          self.saveGeometry()))
        if self.filename is not None:
            settings.setValue("Settings/lastFile", QVariant(self.filename))
        if self.db.isOpen():
            # salva larghezza colonne tabella
            for column in range(1, self.sModel.columnCount()-1):
                width = self.sTableView.columnWidth(column)
                if width:
                    settings.setValue("Settings/sTableView/%s" % column,
                                        QVariant(width))
            self.db.close()
            del self.db

    def setupModels(self):
        """
            Initialize all the application models
        """
        # setup slaveModel
        self.sModel = QSqlTableModel(self)
        self.sModel.setTable(QString("ddtslave"))
        self.sModel.setHeaderData(SID, Qt.Horizontal, QVariant("ID"))
        self.sModel.setHeaderData(SQT, Qt.Horizontal, QVariant("Qt"))
        self.sModel.setHeaderData(SDESC, Qt.Horizontal, QVariant("Descrizione"))
        self.sModel.setHeaderData(SMID, Qt.Horizontal, QVariant("idlegato"))
        self.sModel.setEditStrategy(QSqlTableModel.OnRowChange)
        self.sModel.select()

        # setup masterModel
        self.mModel = QSqlRelationalTableModel(self)
        self.mModel.setTable(QString("ddtmaster"))
        self.mModel.setSort(MDATA, Qt.AscendingOrder)
        self.mModel.setRelation(MIDCLI, QSqlRelation("clienti",
                                            "id", "ragsoc"))
        self.mModel.select()

    def setupMappers(self):
        '''
            Initialize all the application mappers
        '''
        self.mapper = QDataWidgetMapper(self)
        self.mapper.setSubmitPolicy(QDataWidgetMapper.ManualSubmit)
        self.mapper.setModel(self.mModel)
        self.mapper.setItemDelegate(QSqlRelationalDelegate(self))
        self.mapper.addMapping(self.dateEdit, MDATA)
        self.mapper.addMapping(self.ddtLineEdit, MDDT)
        relationModel = self.mModel.relationModel(MIDCLI)
        relationModel.setSort(CRAGSOC, Qt.AscendingOrder)
        relationModel.select()
        self.cliComboBox.setModel(relationModel)
        self.cliComboBox.setModelColumn(relationModel.fieldIndex("ragsoc"))
        self.mapper.addMapping(self.cliComboBox, MIDCLI)
        self.mapper.addMapping(self.cauLineEdit, MCAU)
        self.mapper.addMapping(self.noteLineEdit, MNOTE)
        self.mapper.toFirst()

    def setupTables(self):
        """
            Initialize all the application tablesview
        """
        self.sTableView.setModel(self.sModel)
        self.sTableView.setColumnHidden(SID, True)
        self.sTableView.setColumnHidden(SMID, True)
        self.sTableView.setWordWrap(True)
        self.sTableView.resizeRowsToContents()
        self.sTableView.setAlternatingRowColors(True)
        self.sItmSelModel = QItemSelectionModel(self.sModel)
        self.sTableView.setSelectionModel(self.sItmSelModel)
        self.sTableView.setSelectionBehavior(QTableView.SelectRows)
        self.sTableView.setTabKeyNavigation(False)
        self.myDelegate = MyQSqlRelationalDelegate(self)
        self.sTableView.setItemDelegate(self.myDelegate)
        self.connect(self.myDelegate, SIGNAL("addDettRecord()"),
            self.addDettRecord)

    def setupUiSignals(self):
        self.connect(self.printPushButton, SIGNAL("clicked()"),
                    self.printDdt)
        self.connect(self.addMPushButton, SIGNAL("clicked()"),
                    self.addDdtRecord)
        self.connect(self.delMPushButton, SIGNAL("clicked()"),
                    self.delDdtRecord)
        self.connect(self.addSPushButton, SIGNAL("clicked()"),
                    self.addDettRecord)
        self.connect(self.delSPushButton, SIGNAL("clicked()"),
                    self.delDettRecord)
        self.connect(self.firstMPushButton, SIGNAL("clicked()"),
                    lambda: self.saveRecord(MainWindow.FIRST))
        self.connect(self.prevMPushButton, SIGNAL("clicked()"),
                    lambda: self.saveRecord(MainWindow.PREV))
        self.connect(self.nextMPushButton, SIGNAL("clicked()"),
                    lambda: self.saveRecord(MainWindow.NEXT))
        self.connect(self.lastMPushButton, SIGNAL("clicked()"),
                    lambda: self.saveRecord(MainWindow.LAST))

    def saveRecord(self, where):
        if not self.db.isOpen():
            self.statusbar.showMessage(
                "Database non aperto...",
                5000)
            return
        row = self.mapper.currentIndex()
        self.mapper.submit()
        self.sModel.revertAll()
        if where == MainWindow.FIRST:
            row=0
        elif where == MainWindow.PREV:
            row = 0 if row <= 1 else row - 1
        elif where == MainWindow.NEXT:
            row += 1
            if row >= self.mModel.rowCount():
                row = self.mModel.rowCount() -1
        elif where == MainWindow.LAST:
            row = self.mModel.rowCount()- 1
        #print("row: %d" % row)
        self.mapper.setCurrentIndex(row)
        self.mmUpdate()

    def printDdt(self):
        '''
            Print Inventory
        '''
        if not self.db.isOpen():
            self.statusbar.showMessage(
                "Database non aperto...",
                5000)
            return

        def makeDDT(copia="Copia Cliente"):
            qmaster = QSqlQuery()
            qcli = QSqlQuery()
            qslave = QSqlQuery()

            curidx = self.mapper.currentIndex()
            currec = self.mModel.record(curidx)
            masterid = currec.value("id").toInt()[0]

            qmaster.exec_("SELECT id,data,ddt,idcli,causale,note "
                        "FROM ddtmaster WHERE ddt = '%s'" % (currec.value("ddt").toString()))
            qmaster.next()
            curcli = qmaster.value(3).toInt()[0]
            qcli.exec_("SELECT id,ragsoc,indirizzo,piva "
                        "FROM clienti WHERE id = %d" % (curcli))
            qslave.exec_("SELECT mmid,qt,desc "
                            "FROM ddtslave WHERE mmid = %s" % (masterid))

            qcli.next()
            # variabili utili alla stampa del report
            dataddt = currec.value("data").toDate().toString(DATEFORMAT)
            causaleddt = currec.value("causale").toString()
            noteddt = currec.value("note").toString()
            numddt = currec.value("ddt").toString()
            cliragsoc = qcli.value(1).toString()
            cliind = qcli.value(2).toString()
            clipiva = qcli.value(3).toString()

            from reportlab.pdfgen.canvas import Canvas
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib.units import cm
            from reportlab.lib.enums import TA_LEFT,TA_RIGHT,TA_CENTER
            from reportlab.platypus import Spacer, SimpleDocTemplate, Table
            from reportlab.platypus import TableStyle, Paragraph, KeepTogether
            from reportlab.rl_config import defaultPageSize
            from reportlab.lib import colors
            from reportlab.lib.pagesizes import letter, A4
            from reportlab.lib import textsplit

            PAGE_WIDTH, PAGE_HEIGHT=defaultPageSize
            styles = getSampleStyleSheet()
            styleN = styles['Normal']
            styleH = styles['Heading1']
            styleH.alignment=TA_CENTER
            Elements = []
            #add some flowables
            p=Paragraph
            ps=ParagraphStyle

            Author = "Stefano Zamprogno"
            URL = "http://www.zamprogno.it/"
            email = "time@zamprogno.it"

            pageinfo = "%s / %s" % (Author, email)

            def myLaterPages(c, doc):
                c.saveState()
                c.setFont("Times-Bold", 82)
                c.rotate(45)
                c.setFillColorRGB(0.9,0.9,0.9)
                c.drawString(11*cm,2*cm, copia)
                c.rotate(-45)
                c.setFillColorRGB(0,0,0)
                # HEADER
                c.setLineWidth(0.1*cm)
                c.setStrokeColorRGB(0,0,0)
                c.line(1.8*cm,6.2*cm,19.5*cm,6.2*cm)
                c.setStrokeColorRGB(0.5,0.5,0.5)
                c.line(1.9*cm,6.1*cm,19.6*cm,6.1*cm)
                # cerchi carta
                c.circle(0.9*cm,6*cm,0.3*cm, fill=1)
                c.circle(0.9*cm,24*cm,0.3*cm, fill=1)
                c.setFont("Times-Bold", 14)
                c.drawCentredString(5*cm, 28*cm, r1)
                c.setFont("Times-Bold", 9)
                c.drawCentredString(5*cm, 27.5*cm, r2)
                c.drawCentredString(5*cm, 27*cm, r3)
                c.drawCentredString(5*cm, 26.5*cm, r4)
                c.drawCentredString(5*cm, 26*cm, r5)
                # numero ddt e descrizione copia
                c.setFont("Times-Bold", 12)
                c.drawCentredString(18*cm, 28*cm, "DDT N: %s" % (numddt))
                c.setFont("Times-Bold", 7)
                c.drawCentredString(18*cm, 27.6*cm, "(%s)" % (copia))
                # Data e causale
                c.setFont("Times-Bold", 10)
                c.drawString(1.8*cm, 25*cm, "Data:")
                c.drawString(1.8*cm, 24.5*cm, "Causale:")
                c.setFont("Times-Roman", 8)
                c.drawString(4*cm, 25*cm, unicode(dataddt))
                c.drawString(4*cm, 24.5*cm, unicode(causaleddt))
                # Cliente
                c.setFont("Times-Bold", 10)
                c.drawString(11*cm, 25*cm, "Destinatario:")
                c.setFont("Times-Roman", 8)
                c.drawCentredString(16*cm, 25*cm, unicode(cliragsoc))
                c.drawCentredString(16*cm, 24.5*cm, unicode(cliind))
                c.drawCentredString(16*cm, 24*cm, unicode(clipiva))
                # FOOTER
                c.setFont("Times-Bold", 10)
                c.setLineWidth(0.01*cm)
                c.drawString(1.8*cm, 5.5*cm, "Note:")
                c.drawString(12*cm, 5.5*cm, "Data inizio trasporto:")
                c.line(15.5*cm,5.4*cm,19*cm,5.4*cm)
                c.drawString(12*cm, 5*cm, "Aspetto dei beni:")
                c.line(15*cm,4.9*cm,19*cm,4.9*cm)
                c.drawString(12*cm, 4.5*cm, "Numero colli:")
                c.line(15*cm,4.4*cm,19*cm,4.4*cm)
                c.drawString(12*cm, 3.8*cm, "Conducente:")
                c.line(15*cm,3.7*cm,19*cm,3.7*cm)
                c.drawString(12*cm, 3*cm, "Destinatario:")
                c.line(15*cm,2.9*cm,19*cm,2.9*cm)
                c.drawString(1.8*cm, 4*cm, "Trasporto a Mezzo:")
                c.line(2.3*cm,3*cm,7*cm,3*cm)
                c.setFont("Times-Roman", 8)
                strt = 5.5*cm
                for i in textsplit.wordSplit(unicode(noteddt),6*cm,
                                            "Times-Roman", 8):
                    c.drawString(3*cm, strt, i[1])
                    strt -= 0.5*cm
                # note pie' pagina
                c.setFont('Times-Roman',9)
                c.drawString(12.4*cm, 1.5*cm, "Pagina %d %s" % (doc.page, pageinfo))
                c.restoreState()

            # crea il body del ddt
            data = [['Qt', 'Descrizione dei beni, natura e qualità'],]
            tot = 0
            while qslave.next():
                tot += qslave.value(1).toInt()[0]
                data.append([qslave.value(1).toInt()[0],
                            p(unicode(qslave.value(2).toString()),
                                ps(name='Normal'))])

            Elements.append(Table(data,colWidths=(1*cm,16.5*cm),repeatRows=1,
                                style=(
                                        ['LINEBELOW', (0,0), (-1,0),
                                            1, colors.black],
                                        ['BACKGROUND',(0,0),(-1,0),
                                            colors.lightgrey],
                                        ['GRID',(0,0),(-1,-1), 0.2,
                                            colors.black],
                                        ['FONT', (0, 0), (-1, 0),
                                            'Helvetica-Bold', 10],
                                        ['VALIGN', (0,0), (-1,-1), 'TOP'],

                                )))

            summary = []
            summary.append(Spacer(0.5*cm, 0.5*cm))
            summary.append(Paragraph("<para align=right><b>___________________________________"
                            "</b></para>", styleN))
            summary.append(Paragraph("<para align=right><b>TOTALE ARTICOLI: "
                            "%d</b></para>" % tot, styleN))

            Elements.append(KeepTogether(summary))

            # 'depure' numddt
            numddt = numddt.replace("/",".")

            doc = SimpleDocTemplate(os.path.join(os.path.dirname(__file__),
                            "ddt%s.%s.pdf" % (numddt, copia)),topMargin=6*cm, bottomMargin=6*cm)
            doc.build(Elements,onFirstPage=myLaterPages,onLaterPages=myLaterPages)

            subprocess.Popen(['gnome-open',os.path.join(os.path.dirname(__file__),
                            "ddt%s.%s.pdf" % (numddt, copia))])

        if self.copiaCliCheckBox.isChecked():
            makeDDT()
        if self.copiaIntCheckBox.isChecked():
            makeDDT(copia="Copia Interna")
        if self.copiaVettCheckBox.isChecked():
            makeDDT(copia="Copia Vettore")

def main():
    app = QApplication(sys.argv)
    app.setOrganizationName(DDTORG)
    app.setOrganizationDomain(DDTDOMAIN)
    app.setApplicationName(DDTAPP)

    form = MainWindow()
    form.show()
    form.raise_()
    app.exec_()
    del form

main()
