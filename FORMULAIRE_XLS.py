from tkinter import *
from tkcalendar import *
from tkinter import ttk, messagebox
import tkinter as tk
import openpyxl,xlrd
from openpyxl.workbook import Workbook

import pathlib

class FormulaireExcel :
    def __init__(self, root):
        self.root=root
        self.root.title("formulaire avec Excel")
        self.root.geometry("575x575+300+100")

        frame1 = Frame(self.root, bg="blue")
        frame1.place(x=40, y=50, height=500,width=500)


        title= Label(frame1,text=" Formulaire ",font=( "Algerian",20,"bold") , bg= "blue" , fg="white").place ( x= 50 , y= 30 )

        text_prenom= Label(frame1,text=" Prènom ",font=( "Algerian",10,"bold") , bg= "blue" , fg="white").place ( x= 50 , y= 100 )
        self.ecri_prenom= Entry( frame1, font=("time new roman", 15), bg= "White")
        self.ecri_prenom.place( x= 180 , y= 100)
        text_nom= Label(frame1,text=" Nom ",font=( "Algerian",10,"bold") , bg= "blue" , fg="white").place ( x= 50 , y= 140 )
        self.ecri_nom= Entry( frame1, font=("time new roman", 15), bg= "White")
        self.ecri_nom.place( x= 180 , y= 140)
        text_Email= Label(frame1,text=" E-mail ",font=( "Algerian",10,"bold") , bg= "blue" , fg="white").place ( x= 50 , y= 180 )
        self.ecri_Email= Entry( frame1, font=("time new roman", 15), bg= "White")
        self.ecri_Email.place( x= 180 , y= 180)
        text_Password= Label(frame1,text=" Password ",font=( "Algerian",10,"bold") , bg= "blue" , fg="white").place ( x= 50 , y= 220 )
        self.ecri_Password= Entry( frame1, font=("time new roman", 15), bg= "White")
        self.ecri_Password.place( x= 180 , y= 220)
        text_confPassword= Label(frame1,text=" confirm Password ",font=( "Algerian",10,"bold") , bg= "blue" , fg="white").place ( x= 50 , y=260)
        self.ecri_confPasword= Entry( frame1, font=("time new roman", 15), bg= "White")
        self.ecri_confPasword.place( x= 180 , y= 260)
        text_sexe= Label(frame1,text=" sex ",font=( "Algerian",10,"bold") , bg= "blue" , fg="white").place ( x= 50 , y= 300 )
        self.ecri_sex= ttk.Combobox(  frame1,state=" readonly",  font=("time new roman", 15))
        self.ecri_sex["values"] = (" Homme", "Femme")
        self.ecri_sex.set("Selection")
        self.ecri_sex.place( x = 180 , y = 300)

        self.ecri_sex.place( x= 180 , y= 300)
        text_dateNaissance= Label(frame1,text=" date de naissance ",font=( "Algerian",10,"bold") , bg= "blue" , fg="white").place ( x= 50 , y= 340 )
        self.ecri_dateNaissance= DateEntry( frame1, font=("time new roman", 15), bg= "White", date_pattern = "dd/mm/yy")
        self.ecri_dateNaissance.place( x= 180 , y= " 340")

        b1 = Button(frame1,text="Valider",font=( "new roman",15) , bg= "limegreen",bd= 5).place(x = 80 , y= 400, width= 150)
        b2 = Button(frame1,text="Reinitialiser",font=( "new roman",15) , bg= "limegreen",bd= 5).place(x =   300 , y= 400, width= 150)

        ####3 CRER UN FICHIER EXCEL
        fichier = pathlib.Path( r"C:\Users\HOUCINE\tunisians-loves-programming\FORMILAIRE EXCEL\Base_Donne.xlsx")
        if fichier.exists():
            pass
        else:
            fichier = Workbook ()
            sheet = fichier.active
            sheet ["A1"] = "Prènom"
            sheet["B1"] =  " Nom"
            sheet["C1"] =  " E-mail"
            sheet["D1"] = " Password "
            sheet["E1"] = " sex "
            sheet["F1"] = " Date de naissance"

            fichier.save( r"C:\Users\HOUCINE\tunisians-loves-programming\FORMILAIRE EXCEL\Base_Donne.xlsx")



root=tk.Tk()
obj=FormulaireExcel(root)
root.mainloop()
