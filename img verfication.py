#!/usr/bin/env python
# coding: utf-8

# In[37]:


from PIL import Image   # upload image using PIL
import pandas as pd     
import pytesseract as pt # EXTRACT image to string using PYTESSRACT

img1 = pt.image_to_string(Image.open("result_Page_1.jpg"))

img2 = pt.image_to_string(Image.open("result_Page_2.jpg"))

img3 = pt.image_to_string(Image.open("result_Page_3.jpg"))

img4 = pt.image_to_string(Image.open("result_Page_4.jpg"))

img5 = pt.image_to_string(Image.open("result_Page_5.jpg"))

img6 = pt.image_to_string(Image.open("result_Page_6.jpg"))

img7 = pt.image_to_string(Image.open("result_Page_7.jpg"))

# img extract using pytesseract engine

print("IMAGE1  :-   " +  img1)

print("IMAGE2  :-   " + img2)

print("IMAGE3  :-   " + img3)

print("IMAGE4  :-   " + img4)

print("IMAGE5  :-   " + img5)

print("IMAGE6  :-   " + img6)

print("IMAGE7  :-   " + img7)
        


# In[38]:


from openpyxl import Workbook  # to create a excel sheet using this library

LOT = ["Device Name","REF NML","LOT","SYMBOL"]     # creating a default column name 
d1,r1,l1,s1 = LOT

wb = Workbook()

wb

sheet = wb.active
sheet["A1"] = d1
sheet["B1"] = r1
sheet["C1"] = l1
sheet["D1"] = s1


wb.save(filename = "C:\\Users\\jeeva\\Downloads\\assignment_img_vrf\\imgc.xlsx")


# In[57]:


# img 1 output : 

sy1 = pt.image_to_string(Image.open("1.png"))
sy2 = pt.image_to_string(Image.open("2.png"))
sy3 = pt.image_to_string(Image.open("3.png"))
sy6 = pt.image_to_string(Image.open("6.png"))
sy4 = pt.image_to_string(Image.open("4.png"))
sy5 = pt.image_to_string(Image.open("5.png"))
sy7 = pt.image_to_string(Image.open("7.png"))
sy8 = pt.image_to_string(Image.open("8.png"))
sy9 = pt.image_to_string(Image.open("9.png"))

def withdevice1(dn):
    
    if dn1 in img1:
        print( dn  +  " Pulse Oximeter")
    else:
        print(0)
        
dn1 = "Pulse Oximeter"        
dn0 = 0

withdevice1("Device Name :") 
            
def withdevice1(rn):
    
    if rn1 in img1:
        print( rn +" 903055")
    else:
        print(0)
        
rn1 = "903055"    
rn0 = 0

withdevice1("REF NML:") 

def withdevice1(lt):
    
        if lo1 in img1:
            print( lt + " 34683")
        else:
            print(lo0)
    
lo1 = "34683"
lo0 = 0
    
withdevice1("LOT: ")
       
def withdevice1(sy):
    if sy2 in img1:
        if sy3 in img1:
            if sy6 in img1:
                if sy9 in img1:
                    print( sy + " 2369")
                    return  "2369"    
x = withdevice1("Symbol : ")       

sheet["A2"] = dn1
sheet["B2"] = rn1
sheet["C2"] = lo1
sheet["D2"] = x
    
print(values)
    
wb.save(filename = "C:\\Users\\jeeva\\Downloads\\assignment_img_vrf\\imgc.xlsx")    


# In[56]:


# img 2 output : 


def withdevice2(dn):
    
    if dn2 in img2:
        print( dn  +  " Blood Warmer")
    else:
        print(dn20)

dn2 = "Blood Warmer"
dn20 = 0
        
withdevice2("Device Name :")     
            
def withdevice2(rn):
    
    if rn2 in img2:
        print( rn +" 903090")
    else:
        print(rn20)
        
rn2 = "903090"
rn20 = 0        

withdevice2("REF NML:")   

def withdevice2(lt):
    
    if lt2 in img2:
            print( lt + " 34641")
    else:
            print(lo20)
        
lt2 =  "34641"    
lo20 = 0
      
withdevice2("LOT: ")

def withdevice2(sy):
    if sy1 in img2:
        if sy5 in img2:
            if sy7 in img2:
                print( sy + " 157")
                return "157"  
    
x = withdevice2("Symbol : ")                   

sheet["A3"] = dn2
sheet["B3"] = rn2
sheet["C3"] = lo2
sheet["D3"] = x
        
wb.save(filename = "C:\\Users\\jeeva\\Downloads\\assignment_img_vrf\\imgc.xlsx")    


# In[55]:


# img 3 output : 

def withdevice3(dn):
    
    if dn3 in img3:
        print( dn  +  "  C-Pap Machine")
    else:
        print(0)
        
dn3 = "C-Pap Machine"
dn30 = 0
        
withdevice3("Device Name :") 
    
def withdevice3(rn):
    
    if rn3 in img3:
        print( rn +" 903105")
    else:
        print(rn30)
        
rn3 = "903105"
rn30 = 0
   
withdevice3("REF NML:") 

def withdevice3(lt):
    
    if lo3 in img3:
        print( lt + " 34662")
    else:
        print(lo30)

lo3 = "34662"
lo30 =  0

withdevice3("LOT: ")

def withdevice3(sy):
    if sy1 in img3:
        if sy5 in img3:
            if sy7 in img3:
                if sy8 in img3:
                    if sy9 in img3:
                        print( sy + " 15789")
                        return "15789"
x = withdevice3("Symbol : ")

sheet["A4"] = dn3
sheet["B4"] = rn3
sheet["C4"] = lo3
sheet["D4"] = x
        
wb.save(filename = "C:\\Users\\jeeva\\Downloads\\assignment_img_vrf\\imgc.xlsx")    


# In[54]:


# img 4 output : 


def withdevice4(dn):
    
    if "ECG" in img4:
        print( dn  +  "  ECG Machine")
    else:
        print(0)
        
dn4 = "ECG"
dn40 = 0
            
withdevice4("Device Name :")     
    
def withdevice4(rn):
    
    if rn4 in img4:
        print( rn +" 903060")
    else:
        print(0)
        
rn4 = "903060"
rn40 = 0

withdevice4("REF NML:")  

def withdevice4(lt):
    
    if lo4 in img4:
        print( lt + " 34690")
    else:
        print(0)
        
lo4 = "34690"        
lo40 = 0        
  
withdevice4("LOT: ")

def withdevice4(sy):
    if sy2 in img4:
        if sy5 in img4:
            if sy7 in img4:
                if sy8 in img4:
                    print( sy + " 2578")
                    return "2578"
x = withdevice4("Symbol : ")       

sheet["A5"] = dn4
sheet["B5"] = rn4
sheet["C5"] = lo4
sheet["D5"] = x
        
wb.save(filename = "C:\\Users\\jeeva\\Downloads\\assignment_img_vrf\\imgc.xlsx")    


# In[53]:


# img 5 output : 


def withdevice5(dn):
    
    if dn5 in img5:
        print( dn  +  "  HFNC Machine")
    else:
        print(0)
        
dn5 = "HFNC" 
dn50 = 0
        
withdevice5("Device Name :")     
            
def withdevice5(rn):
    
    if rn5 in img5:
        print( rn +" 903095")
    else:
        print(0)
        
rn5 = "903095"
rn50 = 0

withdevice5("REF NML:") 

def withdevice5(lt):
    
    if lo5 in img5:
        print( lt + " 34648")
    else:
        print(0)
        
lo5 =  "34648"
lo50 = 0
   
withdevice5("LOT: ")

def withdevice5(sy):
    if sy2 in img5:
        if sy5 in img5:
            if sy7 in img5:
                if sy8 in img5:
                    print( sy + " 2578")
                    return "2578"
x = withdevice5("Symbol : ")       

sheet["A6"] = dn5
sheet["B6"] = rn5
sheet["C6"] = lo5
sheet["D6"] = x
        
wb.save(filename = "C:\\Users\\jeeva\\Downloads\\assignment_img_vrf\\imgc.xlsx")    


# In[52]:


# img 6 output : 


def withdevice6(dn):
    
    if dn6 in img6:
        print( dn  +  "  Infusion Pump")
    else:
        print(0)
        
dn6 = "Pump"
dn60 = 0        

withdevice6("Device Name :")

def withdevice6(rn):
    
    if rn6 in img6:
        print( rn +" 903065")
    else:
        print(0)
        
rn6 = "903065"
rn60 = 0
     
withdevice6("REF NML:")   

def withdevice6(lt):
    
    if lo6 in img6:
        print( lt + " 34697")
    else:
        print(0)
    
lo6 =  "34697"
lo60 = 0
 
withdevice6("LOT: ")


def withdevice6(sy):
    if sy2 in img6:
        if sy5 in img6:
            if sy9 in img6:
                if sy8 in img6:
                    print( sy + " 2589")
                    return "2589"
x = withdevice6("Symbol : ")       

sheet["A7"] = dn6
sheet["B7"] = rn6
sheet["C7"] = lo6
sheet["D7"] = x
        
wb.save(filename = "C:\\Users\\jeeva\\Downloads\\assignment_img_vrf\\imgc.xlsx")    


# In[51]:


# img 7 output : 


def withdevice7(dn):
    
    if dn7 in img7:
        print( dn  +  "  NIBP Machine")
    else:
        print(0)
        
dn7 =  "NIBP"       
dn70 = 0        

withdevice7("Device Name :")   

def withdevice7(rn):
    
    if rn7 in img7:
        print( rn +" 903050")
    else:
        print(0)
        
rn7 =  "903050"      
rn70 = 0    
  
withdevice7("REF NML:")

def withdevice7(lt):
    
    if lo7 in img7:
        print( lt + " 34676")
    else:
        print(0)
        
lo7 = "34676"
lo70 = 0

withdevice7("LOT: ")

def withdevice7(sy):
    if sy1 in img7:
        if sy2 in img7:
            if sy7 in img7:
                if sy8 in img7:
                    print( sy + " 127")
                    return " 127"   
                        
x = withdevice7("Symbol : ")       

sheet["A8"] = dn7
sheet["B8"] = rn7
sheet["C8"] = lo7
sheet["D8"] = x
        
wb.save(filename = "C:\\Users\\jeeva\\Downloads\\assignment_img_vrf\\imgc.xlsx") 


# In[ ]:




