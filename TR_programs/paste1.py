def paste(ws,book,Sheet,source,Graph_Tittle):
        import xlsxwriter
        import numpy as np
        from TR_programs import graph1,paste1
        def isfloat(string):
            try:
                float(string)
                return True
            except ValueError:
                return False




        f=open(source,'r')
        data=f.read()
        data1=np.array(data.split('\n'))
        f.close()
        okp=0
        k=data1.shape[0]
        for i in range(1, k + 1):
            H=data1[i-1]
            matrix2 = np.array(H.split())
            for j in range(1, matrix2.shape[0] + 1):
                matrix=matrix2[j-1]
                if not matrix.isdecimal():
                    if isfloat(matrix):
                        matrix=float(matrix)
                K=str(matrix)
                if K.isdecimal():
                    matrix=int(matrix)

                ws.write(i,j-1,matrix)
                if i==1 and j==1:
                    global min
                    min=matrix
                if i==k-1 and j==1:
                    global max
                    max=matrix

        graph1.add_to_chart(ws,book,Sheet,Graph_Tittle,data1.shape[0]-1,min,max)
        graph1.add_to_chart1(ws,book,Sheet,Graph_Tittle,data1.shape[0]-1,min,max)
        del data1, data
