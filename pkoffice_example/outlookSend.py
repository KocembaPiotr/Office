from pkoffice import outlook
import pandas as pd
import numpy as np

file_tmp = outlook.ATTACHMENTS_TMP_PATH + 'text.xlsx'

df = pd.DataFrame({'ID': np.arange(1, 100, 1)})
df.to_excel(file_tmp, index=False, sheet_name='Data')

outlook.send_mail('KocembaPiotr@gmail.com', '', 'Test', 'Test',
                  attachments_list=[file_tmp])

outlook.delete_files([file_tmp])
