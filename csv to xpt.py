import xport
import xport.v56
import pandas as pd

df = pd.DataFrame({
    'alpha': [10, 20, 30],
    'beta': ['x', 'y', 'z'],
})
ds = xport.Dataset(df, name='MAX8CHRS')
with open(r'K:\mashuaifei\xpt\example2.xpt', 'wb') as f:
    xport.v56.dump(ds, f)
