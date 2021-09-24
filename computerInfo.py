import wmi
import pymysql.cursors
import time

c=wmi.WMI()

config = {
    'host':'mysql ip',
    'port':3306,
    'user':'mysql user id',
    'password':'mysql user password',
    'db':'mysql database',
    'charset':'utf8mb4',
    'cursorclass':pymysql.cursors.DictCursor,
}

connection = pymysql.connect(**config)

cursor = connection.cursor()


#電腦資訊
for Computer in c.Win32_ComputerSystem():
    Computer_Caption = Computer.Caption
    Computer_UserName = Computer.UserName

#作業系統
for os in c.Win32_OperatingSystem():
    os_System =os.Caption + " " + os.OSArchitecture
    
#主機板的資訊
bios = c.Win32_BIOS()[0]
bios_Manfacturer = bios.Manufacturer
bios_BIOSVersion = bios.BIOSVersion
bios_Name = bios.Name
bios_SerialNumber = bios.SerialNumber

#CPU資訊
cpuArr = c.Win32_Processor()
for cpu in cpuArr:
    cpu_Name = cpu.Name
    cpu_Manufacturer = cpu.Manufacturer
    cpu_AddressWidth = cpu.AddressWidth
    cpu_Caption = cpu.Caption

#記憶體資訊
memoryArr = c.Win32_PhysicalMemory()
memoryManufacturer=[]
memoryCapacity=[]
memoryPartNumber=[]
memoryAllList=[]
memoryList={}
for memory in memoryArr:
    memoryManufacturer.append(memory.Manufacturer)
    memoryCapacity.append("%.fGB" % (int(memory.Capacity)/1073741824 ))
    memoryPartNumber.append(memory.PartNumber)
    memoryList={'Capacity':memoryCapacity,'Manufacturer':memoryManufacturer,'PartNumber':memoryPartNumber}
memory_amount = len(memoryList['Capacity'])    
memoryListStr = str(memoryList)

#磁碟機資訊 
hddModel=[]
hddSize=[]
hddList={}
for hdd in c.Win32_DiskDrive():
    hddModel.append(hdd.Model)
    hddSize.append("%.fGB" % (int(hdd.Size)/1073741824))
    hddList={'Model':hddModel,'Size':hddSize}
hdd_amount = len(hddList['Model'])
hddListStr = str(hddList)

IpList = []
for network in c.Win32_NetworkAdapterConfiguration(IPEnabled=1):
    ip_Address = network.IPAddress[0]
    ip_MAC = network.MACAddress
    IpList.append(ip_Address)
    IpList.append(ip_MAC)
ip_Address = IpList[0]
ip_MAC = IpList[1]

# print('安裝軟體:')
softwareList=[]
for p in c.Win32_Product():
    software = p.Caption, p.Version
    softwareadd = software[0]," 版本", software[1]
    softwareList.append(softwareadd)
softwareListStr = str(softwareList)

nowTime = time.localtime()
updateTime = time.strftime("%Y/%m/%d-%H:%M:%S",nowTime)

sqlSelectUserName= "SELECT Computer_Caption FROM userComputer WHERE Computer_Caption = '%s' " % (Computer_Caption)
sqlAdd = "INSERT INTO `userComputer`(`Computer_Caption`,`Computer_UserName`,`Os_System`,`Bios_Manfacturer`,`Bios_Name`,`Bios_SerialNumber`,`Cpu_Name`,`Cpu_Manufacturer`,`Cpu_AddressWidth`,`Cpu_Caption`,`Memory_amount`,`MemoryList`,`Hdd_amount`,`HddList`,`ip_Address`,`ip_MAC`,`Software`,`UpdateTime`) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"

cursor.execute(sqlSelectUserName)
try:
   userSelectData =cursor.fetchall()
   if userSelectData[0]["Computer_Caption"] == Computer_Caption:
         sqlUpdateTime = "UPDATE userComputer SET Os_System=%s,Bios_Manfacturer=%s,Bios_Name=%s,Bios_SerialNumber=%s,Cpu_Name=%s,Cpu_Manufacturer=%s,Cpu_AddressWidth=%s,Cpu_Caption=%s,Memory_amount=%s,MemoryList=%s,Hdd_amount=%s,HddList=%s,ip_Address=%s,ip_MAC=%s,Software=%s,UpdateTime=%s WHERE Computer_Caption=%s"
         cursor.execute(sqlUpdateTime,(os_System,bios_Manfacturer,bios_Name,bios_SerialNumber,cpu_Name,cpu_Manufacturer,cpu_AddressWidth,cpu_Caption,memory_amount,memoryListStr,hdd_amount,hddListStr,ip_Address,ip_MAC,softwareListStr,updateTime,Computer_Caption))
except IndexError:
    cursor.execute(sqlAdd,(Computer_Caption,Computer_UserName,os_System,bios_Manfacturer,bios_Name,bios_SerialNumber,cpu_Name,cpu_Manufacturer,cpu_AddressWidth,cpu_Caption,memory_amount,memoryListStr,hdd_amount,hddListStr,ip_Address,ip_MAC,softwareListStr,updateTime))

connection.commit()

