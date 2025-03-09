def H_I(listA):
    value = 0
    for x in listA:
        value += (((x - 4) / (2 * 4)) + 0.5)
    return value / (len(listA))

def  normalize_decision_matrix(listA):
    for y in range(len(listA[0])):
        sumx=0
        for x in range(len(listA)):
            sumx+=listA[x][y]
        for x in range(len(listA)):
            listA[x][y]=listA[x][y]/sumx
        print(listA)


environmental1_social=[]
environmental1_economic=[]
social1_social=[]
economic=[]
a =[[.60,3020],[0.76,3050],[.64,3270]]

normalize_decision_matrix(a)