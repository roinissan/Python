import pypyodbc,sys,logging,os
from datetime import datetime
from datetime import timedelta

class AccessToSql:
    def __init__(self):
        self.dateFormat = ""
        self.ID = ""
        self.accessCursor = None
        self.dataFilePath = ""
        self.data = None

    def configLogging(self,path):
        loggingDirectoryPath = os.path.join(os.path.dirname(path), "logs/log.txt")
        if not os.path.exists(os.path.dirname(loggingDirectoryPath)):
            os.makedirs(os.path.dirname(loggingDirectoryPath))
        logging.basicConfig(filename=loggingDirectoryPath, level=logging.WARNING, format="%(asctime)s:%(levelname)s:%(message)s")

    def setCursor(self,connectionParameters):
        connection = pypyodbc.win_connect_mdb(connectionParameters)
        self.accessCursor = connection.cursor()

    def dataValidation(self,accessTableName,satecModelColNames,sqlColumns):
        if (len(satecModelColNames) != len(sqlColumns) - 1):
            logging.critical("Satec {}:The SQL Table has different number of attributes compared to the access data table ({})".format(self.ID,self.dataFilePath))
            sys.exit(0)

        tableValidID = self.validTable(self.accessCursor, accessTableName)
        if (not tableValidID):
            logging.critical("Satec {}:Could not find the data table - {} in file".format(self.ID,accessTableName))
            sys.exit(0)

        colNames = self.validTableCol(accessTableName[tableValidID],satecModelColNames)
        index = 0
        for colName in colNames[0]:
            if (colName == -1):
                logging.critical("Satec %i:Missing column in access data table - %s" % (self.ID, accessTableName, satecModelColNames[index]))
                sys.exit(0)
            index = index + 1
        return colNames

    def validTable(self,odbcCurser, tableNames):
        tableID = 0
        for table in tableNames:
            tableName = table.lower()
            for table in odbcCurser.tables():
                if (tableName == table[2].lower()):
                    return tableID
            tableID = tableID +  1
        return 0

    def validTableCol(self, accessTableName, satecModelColNames):
        trfTable = [-1] * len(satecModelColNames)
        query = "select * from [%s]" % (accessTableName)
        queryTable = self.accessCursor.execute(query)
        dateCol = 0
        for index in range(len(satecModelColNames)):
            for accessCol in queryTable.description:
                accessColName = accessCol[0]
                if (satecModelColNames[index].lower() in accessColName):
                    trfTable[index] = accessColName
                    if ("time" in accessColName or "date" in accessColName):
                        dateCol = index
                    break
        return [trfTable, dateCol]

    def deleteEmpty(self,accessDataTable, dateCol):
        cleanDataTable = []
        for row in accessDataTable:
            if (len(row[dateCol].strip()) != 0):
                cleanDataTable.append(row)
        return cleanDataTable

    def setData(self,colNames,accessTableName):
        query = "SELECT "
        for colName in colNames[0]:
            if (colName == -1):
                return
            query = query + "[%s]," % (colName)
        query = query[:-1] + r" FROM [%s]" % (accessTableName)

        self.data = self.accessCursor.execute(query).fetchall()
        self.accessCursor.close()
        self.data = self.deleteEmpty(self.data, colNames[1])


    def identifyTimeFormat(self,time):
        if("." in time):
            return "%H:%M:%S.%f"
        return "%H:%M:%S"

    def identifyDateFormat(self,dateCol):
        seperator = ""
        date = self.data[0][dateCol].split()[0]
        if ("/" in date):
            seperator = "/"
        elif ("-" in date):
            seperator = "/"
        else:
            seperator = "."
        dateFormats = ["%d{}%m{}%y".format(seperator, seperator),
                       "%m{}%d{}%y".format(seperator, seperator),
                       "%y{}%m{}%d".format(seperator, seperator)]
        dateFormatsIndex = 0

        for datetimeIndex in range(len(self.data)):
            date = self.data[datetimeIndex][dateCol].split()[0]
            try:
                datetime.strptime(date, dateFormats[dateFormatsIndex])
            except:
                datetimeIndex = 0
                dateFormatsIndex = dateFormatsIndex + 1
                if(dateFormatsIndex == len(dateFormats)):
                    logging.error("Satec {}:Could not format Date - {}".format(self.ID, date))
                    sys.exit(0)
        return dateFormats[dateFormatsIndex]

    def setDateTimeFormat(self,dateCol):
        timeFormat = self.identifyTimeFormat(self.data[0][dateCol].split()[1])
        dateFormat = self.identifyDateFormat(dateCol)
        self.dateFormat = "%s %s" % (dateFormat,timeFormat)

    def sortAndFilter(self, dateCol, lastModified):
        try:
            lastModifiedDate = datetime.strptime(lastModified, '%d/%m/%Y %H:%M:%S')
            lastModifiedDate = lastModifiedDate + timedelta(days=-2)
        except:
            logging.error("Satec {}:Could not format Last Modified Date - {}" .format(self.ID, lastModified))
            datetimeString = '01/01/16 00:00:00'
            lastModifiedDate = datetime.strptime(datetimeString, '%d/%m/%y %H:%M:%S')

        sortedData = sorted(self.data, key=lambda x: datetime.strptime(x[dateCol].strip(), self.dateFormat),reverse=False)
        filterData = []
        for record in sortedData:
            if (datetime.strptime(record[dateCol].strip(), self.dateFormat) > lastModifiedDate):
                filterData.append(record)
        self.data = filterData

    def createInsertQueries(self, sqlTableName, sqlcolNames, satecID, dateCol, accessColNames):
        queries = []
        rowIndex = 0
        for row in self.data:
            if (rowIndex == 0):
                rowIndex = rowIndex + 1
                continue
            updatedRow = []
            for fieldIndex in range(len(row)):
                if (fieldIndex == dateCol):
                    updatedRow.append(row[fieldIndex])
                else:
                    updatedRow.append(row[fieldIndex] - self.data[rowIndex - 1][fieldIndex])
            query = "INSERT INTO %s (" % (sqlTableName)
            for colName in sqlcolNames:
                query = query + "[%s]," % (colName)
            query = query[:-1] + ") VALUES ("
            for sqlColName in sqlcolNames:
                if ("date" in sqlColName.lower() or "time" in sqlColName.lower()):
                    date = updatedRow[dateCol].strip()
                    date = datetime.strptime(date, self.dateFormat)
                    if(date.minute == 59):
                        date = date.replace(second=59) + timedelta(seconds=1)
                    query = query + "'%s'," % (date.strftime('%Y-%m-%d %H:%M:%S'))
                else:
                    j = 0
                    for accessColName in accessColNames:
                        if (sqlColName in accessColName):
                            query = query + "%s," % (updatedRow[j])
                            break
                        j = j + 1
            rowIndex = rowIndex + 1
            query = query + "%s)" % (satecID)
            queries.append(query)
        return queries

    def writeQueriesToFile(self,queries, filePath):
        try:
            queriesFile = open(filePath, "w")
            for query in queries:
                queriesFile.write(query + "\n")
        except:
            logging.critical("Satec {}:Could not write queries to disk - {}".format(self.ID))
        finally:
            queriesFile.close()

    def main(self):
        pypyodbc.lowercase = True
        argumentsList = []
        for arg in sys.argv[1:]:
            keyValue = arg.split('=')
            if (keyValue[0] == "satecModelColNames" or keyValue[0] == "sqlColumns" or keyValue[0] == "accessTableName"):
                argumentsList.append([keyValue[0], keyValue[1].split(",")])
            else:
                argumentsList.append([keyValue[0], keyValue[1]])

        arguments = dict(argumentsList)
        if("connectionParameters" in arguments):
            connectionParameters = arguments["connectionParameters"]
        else:
            connectionParameters = r'DRIVER=Microsoft Access Driver (*.mdb, *.accdb);DBQ=%s;' % (arguments["filePath"])

        self.ID = arguments["satecID"]
        self.configLogging(arguments["queriesFilePath"])
        self.setCursor(connectionParameters)
        self.dataFilePath = arguments["filePath"]

        colNames = self.dataValidation(arguments["accessTableName"],arguments["satecModelColNames"],arguments["sqlColumns"])
        self.setData(colNames,arguments["accessTableName"])
        self.setDateTimeFormat(colNames[1])
        self.sortAndFilter(colNames[1],arguments["lastModified"])

        self.writeQueriesToFile(
            self.createInsertQueries(arguments["sqlTableName"], arguments["sqlColumns"], arguments["satecID"], colNames[1], arguments["satecModelColNames"]),
            arguments["queriesFilePath"])

if __name__ == "__main__":
    accToSql = AccessToSql()
    accToSql.main()