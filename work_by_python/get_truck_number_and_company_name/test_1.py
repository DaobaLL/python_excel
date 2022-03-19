import numpy as np  


a = "sadfasdf"
 
for b in a:
    print(b)

print(np.__version__)

arr = np.array([[1, 2, 3, 4, 5],[6, 7, 8, 9, 10]])

print(arr)

arr2 = np.array(
    [[[2,3,4],[3,4,5]],
    [[4,5,6],[5,6,7]]]
)
print(arr2)

arr3 = np.array([1,2], ndmin=32)

print(arr3)
print('number if dimensions:', arr.ndim)

arr = np.array(['apple', 'banana', 'cherry'])

print(arr.dtype)