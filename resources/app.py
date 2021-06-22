import shutil
import os


dir = os.path.dirname(os.path.realpath(__file__))
print(dir)

for x in os.walk(dir+'/abc'):
    print(x)

# shutil.copytree(f'{tmp_path}/{file_name}', f'{output_path}/{str(render_id)}', ignore=ig_d)
# if os.path.isdir(dir+'/xyz'):
#     print("111")
#     shutil.copytree(dir+'/abc', dir+'/d')
#     print("222")