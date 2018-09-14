
# coding: utf-8

import xlrd
import xlwt

# # Analyse du fichier Excel
# ## import du fichier

wb = xlrd.open_workbook('VAT_file.xlsx')
sh = wb.sheet_by_name('Sheet1')

print(wb)
print(sh)


# In[3]:


colgroup = sh.col_values(0)
print(colgroup)

headers = sh.row_values(0)
print(headers)


# ## Analyse du contenu

# In[4]:


analyze1 = []
'''liste correspondant à la première colonne : 1 si non vide, 0 si vide'''

for i in range((len(colgroup))):
    if colgroup[i]=="":
        analyze1.append(0)
    else:
        analyze1.append(1)

print(analyze1)
print(len(colgroup))


# In[5]:


analyze2 =[]
colreport = sh.col_values(4)
'''liste correspondant à la 5eme colonne, 1 si non vide, 0 si vide'''

for i in range((len(colreport))):
    if colreport[i]=="":
        analyze2.append(0)
    else:
        analyze2.append(1)

print(analyze2)
print(len(colreport))


# In[6]:


analyze3=[]
var = 0
'''attribution de valeur à chaque groupe titre/x_reports. 1 header'''

for i in range(len(colgroup)-1):
    var = var + analyze1[i]
    analyze3.append(var)
analyze3.append(var)

print(analyze1)
print(analyze2)
print(analyze3)


# In[7]:


analyze4=[]
groupmax = analyze3[-1]
for i in range(groupmax):
    analyze4.append(0)
print(analyze4)

'''dictionnaire de paires (numerogroupe)/(présencevaleur)'''

for i in range(len(colgroup)):
    if analyze2[i]!=0:
        analyze4[analyze3[i]-1]+=1

print(analyze4)


# In[8]:


analyze5 =[]

for i in range(len(analyze4)):
    if analyze4[i]!=0:
        analyze5.append(i+1)

print(analyze5)


# ## Résumé

# In[9]:


print(analyze1) # présence d'un élément dans la première colonne
print(analyze2) # présence d'un élément dans la 5eme colonne
print(analyze3) # répartition à chaque ligne d'un numero de groupe
print(analyze4) # donne le nombre de ligne de la 5eme colonne non vide par groupe
print(analyze5) # liste des groupes à reporter


# # Creation d'un fichier de sortie
# ## Creation du fichier vide

# In[10]:


Excel_output = xlwt.Workbook() # Creation du fichier
print(Excel_output)

new_sheet = Excel_output.add_sheet('Sheet1',True) #second argument : True if overwrite possible
print(new_sheet)


# ## Report des lignes à conserver

# In[11]:


rowmax = len(analyze1) #nombre de lignes à analyser

def ajouter_excel(sheet, liste, ligne):
    for i in range(len(liste)):
        sheet.write(ligne,i,liste[i])

ligne_ecriture = 0

for i in range(rowmax):
    if analyze3[i] in analyze5:
        ajouter_excel(new_sheet, sh.row_values(i),ligne_ecriture)
        ligne_ecriture += 1
        print(sh.row_values(i))


# ## Creation Excel format natif

# In[12]:


# Save object in a native format Excel file
Excel_output.save('Output.xls')

