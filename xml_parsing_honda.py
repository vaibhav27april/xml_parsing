from xml.etree import ElementTree
# from xml.etree.cElementTree as ET
import os
import xlsxwriter
import fnmatch
#filename = 'ACC_file.xml'

#full_filename = os.path.abspath(filename)
#print(full_filename)
#root_data = ElementTree.parse(full_filename).getroot()
#print(root_data.tag)


global_refrence_k=0
local_iterator=0



test_list=[['a','v','ssd','sdf','rt','tyu'],['a','v','ssd','sdf','rt','tyu'],['a','v','ssd','sdf','rt','tyu'],['a','v','ssd','sdf','rt','tyu'],['a','v','ssd','sdf','rt','tyu'],['a','v','ssd','sdf','rt','tyu'],['a','v','ssd','sdf','rt','tyu'],['a','v','ssd','sdf','rt','tyu']]
row=1
#check_di = {'tx': {'key1': 1, 'key2': 2, 'key3': 3} , 'tx2': {'k1': 11, 'k2': 22, 'k3': 33} ,
#'tx3': {'kk1': 111, 'KK2': 222, 'KK3': 333}}

#chek_2 = {'tx': {'tt': 1, 'ww': 2, 'ee': 3} , 'tx1': {'tt': 1, 'ww': 2, 'ee': 3} , 'tx2': {'tt': 1, 'ww': 2, #'ee': 3}}

#final_check = {'key1': check_di,'key2': chek_2}
workbook = xlsxwriter.Workbook('Expenses01.xlsx')
worksheet = workbook.add_worksheet()




def excelupdation():


    worksheet.write('A1','TX/RX')
    worksheet.write('B1','NAME')
    worksheet.write('C1','Message')
    worksheet.write('D1','Length')
    worksheet.write('E1','Startbit')
    worksheet.write('F1','DEFAULT VALUE')
    worksheet.write('G1','Message ID')
    worksheet.write('H1','ID-Format')
    worksheet.write('I1','DLC[Byte]')
    worksheet.write('J1','TX-Method')
    worksheet.write('K1','Cycle Time')
    worksheet.write('L1','Byte Order')
    worksheet.write('M1','Channel')
    return


def Excel_data():
 global row
 global Di
 global local_n
 global Local_list
 global local_k
 global global_refrence_k
 global local_k_last_value
 global local_iterator
 global local_iterator_local

 global_refrence_k=local_k-local_k_last_value

 for i in range(0,global_refrence_k):
     print(Di['TxMessage' + str(local_n)]['Name'])
     print(Di['TxMessage' + str(local_n)]['ID'])
     print(Di['TxMessage' + str(local_n)]['Frametype'])
     print(Di['TxMessage' + str(local_n)]['DLC'])
     #print(Di['TxMessage' + str(local_n)]['Value'])
     print(Local_list[local_iterator_local][0])
     print(Local_list[local_iterator_local][1])
     print(Local_list[local_iterator_local][2])
     worksheet.write_string(row,0,'TX')
     worksheet.write_string(row,1,Local_list[local_iterator_local][0])
     worksheet.write_string(row,2,Di['TxMessage' + str(local_n)]['Name'])
     worksheet.write_string(row,3,Local_list[local_iterator_local][2])
     worksheet.write_string(row,4,Local_list[local_iterator_local][1])
     worksheet.write_string(row,5,'0')
     worksheet.write_string(row,6,Di['TxMessage' + str(local_n)]['ID'])
     worksheet.write_string(row,7,Di['TxMessage' + str(local_n)]['Frametype'])
     worksheet.write_string(row,8,Di['TxMessage' + str(local_n)]['DLC'])
     worksheet.write_string(row,9,'FixedPeriodic')
     worksheet.write_string(row,11,'MFL')

     if'Value' in Di['TxMessage' + str(local_n)]:
      worksheet.write_string(row,10,Di['TxMessage' + str(local_n)]['Value'])
     else:
      worksheet.write_string(row,10,'100')


     if(Di['TxMessage' + str(local_n)]['Frametype']=='CAN Extended'):
         worksheet.write_string(row,12,'2')
     else:
         worksheet.write_string(row,12,'1')


     row=row+1
     local_iterator=local_iterator+1
     local_iterator_local=local_iterator_local+1
 local_k_last_value=local_k
 return


def mainengine():
 local_i=0
 local_j=0
 global local_k
 global global_refrence_k
 local_l=0
 global local_n
 for child in root_data:
     print("**********present in child************")
     #print(str(child.tag) + "  text  " + str(root_data[local_i].text).rstrip())
     print(child.tag + " try " + child.text)
     for childern in child:
        print("*************present_in_childern****************")
        #print(" " + childern.tag)
        print(childern.tag + " try " + str(childern.text).rstrip())
        #print(str(childern.tag) + " text " + str(root_data[local_i][local_j].text).rstrip())

        if(childern.tag=='TxMessage'):
         for more_childern in childern:
            #if(more_childern.tag=='signal'):
             #local_i=local_i+1

            print("*************present in more_childern****************")
            #print(str(childern.tag) + " text " + str(root_data[local_i][local_j][local_k].text).rstrip())
            print(more_childern.tag + " try " + str(more_childern.text))
            Di['TxMessage' + str(local_n)] = Di.get('TxMessage' + str(local_n), {})
            Di['TxMessage' + str(local_n)][more_childern.tag]=str(more_childern.text).rstrip()
            check=0
            count_no=0
            for more_sub_children in more_childern:
             if(more_childern.tag=='Attribute') and (more_sub_children.tag=='Name'):
                 continue
             if(more_childern.tag=='Attribute') and (more_sub_children.tag=='Value'):
                 Di['TxMessage' + str(local_n)] = Di.get('TxMessage' + str(local_n), {})
                 Di['TxMessage' + str(local_n)][more_sub_children.tag]=str(more_sub_children.text).rstrip()
                 continue



             if(check==0):
              Local_list.append([])

             Local_list[local_k].append(str(more_sub_children.text))
             print("*************present in more_sub_childern****************")
             print(more_sub_children.tag + " try " + str(more_sub_children.text).rstrip())
             count_no=count_no+1
             if (count_no==11):
                 local_k=local_k+1

             #if(check==0):



             check=1

         Excel_data()
         local_n=local_n+1

         #local_j=0
 return

excelupdation()

path = os.path.dirname(os.path.abspath(__file__))
listing = os.listdir(path)
for infile in listing:
    print("current file is: " + infile)
    full_filename = str(path + '\\' + infile)
    #if(infile.endswith=='.xml'):
    if fnmatch.fnmatch(infile,'*.xml'):
     root_data = ElementTree.parse(full_filename).getroot()
     Di={}
     Local_list=[]
     local_n=0
     local_k=0
     local_k_last_value=0
     local_iterator_local=0
     mainengine()






#if __name__ == '__main__':
 #excelupdation()
 #main_engine()
 #Excel_data()