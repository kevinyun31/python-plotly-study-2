import numpy as np
from matplotlib import pyplot as plt
import math
 
x = np.arange(0, 3*math.pi, 0.1)
y = np.sin(x)
 
plt.plot(x, y)

plt.savefig('./sin.png')

