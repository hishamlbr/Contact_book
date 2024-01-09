from tkinter import * 
import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap import Style
from tkinter import messagebox
import mysql.connector
from xlwt import Workbook
import os

#/////////////// connect with data base////////
try:
    conn=mysql.connector.connect(
        host="localhost",
        user="root",
        password="",
        database="contacts"
        )
    mycur=conn.cursor()
except mysql.connector.Error as r :
    print(r)      


#======== program Interface ==================
root=Tk()
root.geometry("500x480+350+100")
root.title("Contact Book")
root.resizable(False,False)
root.config(bg="#ffe3be")

style = Style(theme="superhero")

#////////////// Functions ////////////////////////////////////
def add() :
    global value
    try :
        
        name=entry_name.get()
        number=entry_contact.get()
        print(name,":",number)
        for i in range(0,100) :
            value=contact_list.get(i)
            if value==name :
                messagebox.showwarning(title="Warning",message="This name already entered !")    
                break
   
        else :
                query="INSERT INTO contact(name,number) VALUES(%s,%s)"
                values=(name,number)
                mycur.execute(query,values)
                conn.commit()
            
                for i  in range (0,100) :
                    contact_list.insert(tk.END,name)
                    entry_name.delete(0,"end")
                    entry_contact.delete(0,"end")
                    break
  
  
    except mysql.connector.Error as er :
        messagebox.showerror(title="Error",message=er)

def fetch_all_contacts() :
    global name,number,data
    query_insert="SELECT name,number FROM contact"
    mycur.execute(query_insert)
    data=mycur.fetchall()
    for row in data:
        name, number = row
        contact_list.insert(tk.END,name)
           
def select_items () :
    
    item=contact_list.curselection()
    item2=contact_list.get(item[0])
   
def delete():
    try:
        itemm = contact_list.curselection()

        itemm2 = contact_list.get(itemm[0])
        del_query = "DELETE FROM contact WHERE name = %s"
        mycur.execute(del_query, (itemm2,))
        conn.commit()
        contact_list.delete(itemm)
        
    
    except:
        messagebox.showerror(title="Error",message="The item not deleted")

def update_db():
        contact_list.delete(iteme)
        new_name=entry_name.get()
        new_number=entry_contact.get()
        for i in range(0,100) :
            value=contact_list.get(i)
            if value==new_name :
                messagebox.showwarning(title="Warning",message="This name already entered !")    
                break
        else:
                
                query_up="UPDATE contact SET name=%s , number=%s WHERE name=%s"
                mycur.execute(query_up,(new_name,new_number,selected_item))
                conn.commit() 
                entry_name.delete(0,"end")
                entry_contact.delete(0,"end")
                contact_list.insert(iteme_index,new_name)
                btn.pack_forget()
                
def edit():
    global iteme_index,selected_item,iteme,btn
    iteme = contact_list.curselection()

    if not iteme:
        messagebox.showwarning(title="Warning!", message="Select an item!")
        return

    iteme_index = iteme[0]
    selected_item = contact_list.get(iteme_index)

    # Find the corresponding contact number based on the selected name
    for name, number in data:
        if name == selected_item:
            entry_name.delete(0, "end")
            entry_contact.delete(0, "end")
            entry_name.insert(tk.END, name)
            entry_contact.insert(tk.END, number)
            break
    btn=Button(root,text="Update",command=update_db) 
    btn.pack()  
     
def show():
    iteme = contact_list.curselection()
    iteme_index = iteme[0]
    selected_item = contact_list.get(iteme_index)

    rot=Tk()
    rot.geometry("400x80+400+250")
    rot.title("Contact information")
    rot.resizable(FALSE,FALSE)
    for name, number in data:
        if name==selected_item:
            info=(name+" : "+number)
            lbl=Label(rot,text=info,font=("Times New Roman",16))
            lbl.pack()
            break
    else :
        print("NOO")
    btn=Button(rot,text="Ok",width=10,command=rot.destroy)
    btn.pack(padx=10,pady=10)


    rot.mainloop()
      
def search ():
    word=entry_search.get()
    if word :
        for i in range(0,100) :
            value=contact_list.get(i)
            if word==value:
                index = contact_list.get(0, tk.END).index(word)
                contact_list.select_set(index)
                contact_list.see(index)
                break
        else :
            print("no that not")
    else:
        messagebox.showwarning(title="Warning",message="Please Enter a name !")

def export():
    wb=Workbook()
    sheet=wb.add_sheet("Export xls")
    sheet.write(0,0,"Name") 
    sheet.write(0,1,"Number")
    for x in range(1,100):
        for name,number in data:
            sheet.write(x,0,name)
            sheet.write(x,1,number)
            x=x+1    
        break
    path="C:\\Users\\HHSS\\Desktop\\Learn_Python\\Tkinter\\exported_contacts.xls"
    msgg=messagebox.askokcancel(title="Export",message="Are you sure you want replace the current file ?")
    if msgg==True:
        if os.path.exists(path) :
            os.remove(path)
            wb.save(path)            
        else :
            wb.save(path) 
    else :
        return




title=Label(text="Contact Book",fg="blue",font=("Times New Roman",24),bg="orange") # title of application
title.pack(pady=5)
#/////////////////////// Labels and Entries ///////////////////////////
label_name=ttk.LabelFrame(text=" Name   :",padding=10)
label_name.place(x=5,y=60)
entry_name = tk.Entry(label_name,font=16)
entry_name.pack()

label_search=ttk.LabelFrame(text=" Search   :",padding=10)
label_search.place(x=260,y=60)
entry_search = tk.Entry(label_search,font=16)
entry_search.pack()
image=PhotoImage(file="C:\\Users\\HHSS\\Desktop\\Learn_Python\\Tkinter\\se.png")
res=image.subsample(2,2)
bt=Button(root,image=res,compound="top",cursor="target",command=search)
bt.place(x=435,y=85)


label_contact=ttk.LabelFrame(text="Contact :",padding=10)
label_contact.place(x=5,y=140)
entry_contact = tk.Entry(label_contact,font=16)
entry_contact.pack()


#//////////////// list box ////////////////////////////
label_list=ttk.LabelFrame(text="Contacts :",padding=10)
label_list.place(x=260,y=140)
contact_list=Listbox(label_list,bg="blue",width=30,height=14)
contact_list.pack()
btn_ex=Button(root,text="Export xls",width=15,height=2,command=export)
btn_ex.place(x=315,y=415)
fetch_all_contacts() #show all the items in the list Box

# Bind the listbox selection event to the select_items function
contact_list.bind('<<ListboxSelect>>', lambda event: select_items())


#///////////////////////////////////////////////////////////////////////////

#//////////////////////// Buttons //////////////////////////////////////////////////////////
bg=PhotoImage(file="C:\\Users\\HHSS\\Desktop\\Learn_Python\\Tkinter\\manbg.png")
cut=bg.subsample(6,6)
image=Label(root,image=cut)
image.place(x=85,y=210)

add_btn=Button(text="Add",width=15,height=2,command=add)                                                    
add_btn.place(x=5,y=215)                                                              
                                                                                       
del_btn=Button(text="Delete",width=15,height=2,command=delete)                                                  
del_btn.place(x=5,y=265)                                                                 
                                                                                        
edit_btn=Button(text="Edit",width=15,height=2,command=edit)
edit_btn.place(x=5,y=315)

res_btn=Button(text="Show contact",width=15,height=2,command=show)
res_btn.place(x=5,y=365)

exit_btn=Button(text="Exit",width=15,height=2,command=root.destroy)
exit_btn.place(x=5,y=415)
#/////////////////////////////////////////////////////////////////////////////////////////////////////









root.mainloop()
