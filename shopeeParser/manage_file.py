import os
import shutil
from datetime import datetime, timedelta

today = datetime.now().strftime("%m_%d_%Y")
yesterday = datetime.strftime(datetime.now() - timedelta(1), "%m_%d_%Y")


for store_name in ['momo', 'shopee','yahoo']:
    base_path = './' + store_name
    for root, dirs, files in os.walk(base_path):
        for dir in dirs:
            if str(dir) == today or str(dir) == yesterday:
                continue
            try:
                folder_path = base_path + '/' + dir
                shutil.rmtree(folder_path)
                print(folder_path + ' deleted')
            except:
                print(folder_path + ' delete failed')
    
    # log_path  = './' + store_name + '_log.txt'
    # if os.path.exists(log_path):
    #     os.remove(log_path)
    # else:
    #     print(store_name + " log does not exist")
    # file = open(log_path, "w") 
    # file.close() 






        



