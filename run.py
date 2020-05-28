#################################################
__author__ = "Teodor Chetrusca"
__copyright__ = "Mead Informatica SRL. NDA "
__license__ = "GPL"
__version__ = "1.0.0"
__maintainer__ = "Teodor Chetrusca"
__email__ = "t.chetrusca@meadinformatica.it"
__status__ = "Production"

#################################################
#crate exe: pyinstaller  run.py --onefile

import sqlite3
import subprocess
import sys
import time
import traceback

class DB:
    '''
    Classe per la gestione del database, il DB è in formato SQLite3
    '''
    conn = None
    connAPEX = None

    def __init__(self):
        self.conn = sqlite3.connect(officescanDB)
        self.connAPEX = sqlite3.connect(apexDB)

    def get_data_pc(self, pc):
        try:
            sql = ''' SELECT  * FROM MaClientInfo  WHERE Name = ? '''
            cur = self.conn.cursor()
            cur.execute(sql, (pc,))
            names = list(map(lambda x: x[0], cur.description))
            records = cur.fetchone()
            return dict(zip(names, records))
        except self.conn.Error:
            print(traceback.format_exc())

    def update_data(self, a, pc):
        try:
            sql = ''' UPDATE MaClientInfo SET Version = ?, FDE_ID = ?, FDE_Ver = ?, FDE_Status = ?,  FDE_Ex_Status = ?, 
                FDE_Act_Type= ?, FDE_DeployTime = ?, FILE_ARMOR_ID = ?, FILE_ARMOR_Ver = ?, FILE_ARMOR_Status = ?, 
                FILE_ARMOR_Ex_Status= ?, FILE_ARMOR_Act_Type = ?, FILE_ARMOR_DeployTime = ?  WHERE Name = ? '''

            data = (a['Version'], a['FDE_ID'], a['FDE_Ver'], a['FDE_Status'], a['FDE_Ex_Status'], a['FDE_Act_Type'],
                    a['FDE_DeployTime'], a['FILE_ARMOR_ID'], a['FILE_ARMOR_Ver'], a['FILE_ARMOR_Status'],
                    a['FILE_ARMOR_Ex_Status'], a['FILE_ARMOR_Act_Type'], a['FILE_ARMOR_DeployTime'], pc)
            cur = self.connAPEX.cursor()
            cur.execute(sql, data)
            self.connAPEX.commit()
            cur.close()
        except self.conn.Error:
            print(traceback.format_exc())



if __name__ == '__main__':

    if len(sys.argv) < 2:
        print("Inserisci il nome della macchina da migrare")
        sys.exit()

    line = sys.argv[1]
    if len(line.strip()) < 1:
        print("Inserisci il nome della macchina da migrare")
        sys.exit()
    else:
        pc = str(line)
    officescanDB = 'db/tmtb.db'
    #apexDB = 'db/tmtb_apex.db'
    apexDB = 'E:/Program Files (x86)/Trend Micro/Apex One/Addon/TMEE/db/tmtb.db'
    db = DB()

    ll = db.get_data_pc(pc)

    stopstr = 'stop-service "Trend Micro Endpoint Encryption Deployment Tool Service"'
    subprocess.run(['powershell', stopstr], shell=True)
    time.sleep(5)
    db.update_data(ll, pc)
    stopstr = 'start-service "Trend Micro Endpoint Encryption Deployment Tool Service"'
    subprocess.run(['powershell', stopstr], shell=True)
