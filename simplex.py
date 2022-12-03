#Смирнов А.И., 2022
#https://github.com/olekzonder
import numpy as np
import warnings
class Simplex:
    def __init__(self,function,constraints, resources, **kwargs):
        self.function = np.array(function)    #целевая функция
        self.constraints  =  np.array(constraints) #ограничения
        self.resources = np.array(resources) #ресурсы (из равенства ограничений)
        self.type = kwargs['type'] #-> max или -> min
        self.outputList = []
        self.minArr = []
        if 'lessThanList' in kwargs.keys(): 
            lessThanList = kwargs['lessThanList']
            if self.type == 'min':
                constraints = [i for i in range(len(self.function)) if i in lessThanList]
            else:
                constraints = [i for i in range(len(self.function)) if i not in lessThanList]
            self.genMatrix(constraints)
        else:
            self.genMatrix([])

    def genMatrix(self,constraints):
        self.x = [i for i in range(1,len(self.function)+1)] #переменные
        if self.type == 'min':
            self.resources = np.array([-i for i in self.resources])
            for i in range(len(self.constraints)):
                for j in range(len(self.constraints[i])):
                    self.constraints[i][j] = -self.constraints[i][j]
            self.constraints = np.array(self.constraints)
        self.p = np.transpose([self.resources])
        basis = np.zeros([len(self.constraints),len(self.constraints)]) #сведение задачи к задаче ЛП
        self.ck = np.zeros([len(self.constraints),1])
        for i in range(len(basis)):
            if i in constraints:
                basis[i][i]=-1
            else:
                basis[i][i] = 1
        self.b = [i for i in range(len(self.function)+1,len(self.constraints)*2+1)] #базисные переменные
        if self.type == 'max':
            self.rp = [0]+[-i for i in self.function]+[0 for j in range(len(self.constraints))] #частичные суммы
            self.matrix = np.hstack((self.p,np.array(self.constraints),basis))
        if self.type == 'min':
            self.rp = [0]+[i for i in self.function]+[0 for j in range(len(self.constraints))] #частичные суммы
            self.matrix = np.hstack((self.p,np.array(self.constraints),basis))
        self.addOutput()

    def genOutput(self):
        output = {}
        data = []
        header = ['Базис','Ck']
        for i in range(len(self.matrix[0])):
            header.append('P'+str(i))
        data.append(header)
        if len(self.minArr):
            header.append("min")
            output['row'] = self.leadRow
            output['col'] = self.leadCol
        for i in range(len(self.matrix)):
            row = []
            row.append('P'+str(self.b[i]))
            row.append(round(self.ck[i][0],6))
            for j in range(len(self.matrix[0])):
                row.append(round(self.matrix[i][j],6))
            if len(self.minArr):
                row.append(self.minArr[i])
            data.append(row)
        data.append(['','']+[round(i,6) for i in self.rp])
        # for i in data:
        #     print(i)
        output['data'] = data
        return output

    def addOutput(self):
        output = self.genOutput()
        if output != None:
            self.outputList.append(output)

    def optimize(self):
        self.iterate()
        tempArr = []
        for i in self.x:
            if i in self.b:
                tempArr.append(self.outputList[len(self.outputList)-1]['data'][self.b.index(i)+1][2])
            else:
                tempArr.append(0)
        self.simplex = self.outputList
        self.result = tempArr
        return self.simplex

    def getResult(self):
        return {'result':self.result,'fx':self.rp[0]}

    def iterate(self):
        if self.type == 'max':
            if len([i for i in self.rp if str(i)[0] == '-'])==0:
                self.minArr = []
                self.addOutput()
                return
            col = self.rp.index(min(self.rp))
        if self.type == 'min':
            if len([i for i in self.rp if i > 0])==0:
                self.minArr = []
                self.addOutput()
                return
            col = self.rp.index(max(self.rp))
        self.minArr = []
        for i in range(len(self.matrix)):
            try:
               val = self.matrix[i][0]/self.matrix[i][col]
               warnings.filterwarnings("ignore")
               if val <= 0:
                raise ZeroDivisionError
               self.minArr.append(val)
            except ZeroDivisionError:
                self.minArr.append(999999999999) #Абсурдно огромное значение для того, чтобы отсеить случаи с делением на 0 или отрицательные значения
        row =  self.minArr.index(min(self.minArr))
        self.leadCol = col
        self.leadRow = row
        self.minArr = [i if i != 999999999999 else '-' for i in self.minArr]
        self.addOutput()
        self.b[row] = col
        self.ck[row] = self.rp[col]
        leadCross = self.matrix[row][col]
        newmatrix = np.zeros([len(self.matrix),len(self.matrix[row])])
        for i in range(len(self.matrix[row])):
            newmatrix[row][i] = self.matrix[row][i]/leadCross
        for i in range(len(self.matrix)):
            if i == row:
                continue
            cross = self.matrix[i][col]
            for j in range(len(self.matrix[i])):
                val = np.float64(self.matrix[i][j]-newmatrix[row][j]*self.matrix[i][col])
                newmatrix[i][j] = val
        for i in range(len(self.rp)):
            self.rp[i] = self.rp[i] - (self.ck[row][0]*self.matrix[row][i])/leadCross
        self.matrix = newmatrix
        self.iterate()