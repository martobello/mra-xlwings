import distutils.dir_util as du

src = 'C:\\Users\\Mike\\git\\mra-xlwings\\src'
dst = 'C:\\Users\\Mike\\Documents\\Excel\\xlwings'

print('Copying ' + src + ' to ' + dst + '\n')

l = du.copy_tree(src, dst,dry_run=1)

print('Files Copied:')
for f in l:
    print(f)