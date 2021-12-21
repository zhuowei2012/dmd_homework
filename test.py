# -*- coding: utf-8 -*-
import os
import math

#for j in range(2, 83):
#    x=(j-2)%9 + 1
#    y=math.ceil((j-1)/9)
#    print("x:{x},y={y}".format(x=x,y=y))
#    print("+++++++++++++++")

NDM = [100,300,500]
NBP = [100,300,500]
for n_dm in NDM:
    for n_bp in NBP:
        print("x:{x},y={y}".format(x=n_dm,y=n_bp))
