import os
import os.path
import shutil
import sys
import win32wnet

def netcopy(host, source, dest_dir, username=None, password=None, move=False):
    """ Copies files or directories to a remote computer. """
    wnet_connect(host, username, password)
    source = covert_unc(host, source)
    # Pad a backslash to the destination directory if not provided.
    if not dest_dir[len(dest_dir) - 1] == '\\':
        dest_dir = ''.join([dest_dir, '\\'])
    # Create the destination dir if its not there.
    if not os.path.exists(dest_dir):
        os.makedirs(dest_dir)
    else:
        # Create a directory anyway if file exists so as to raise an error.
         if not os.path.isdir(dest_dir):
             os.makedirs(dest_dir)
    if move:
        shutil.move(source, dest_dir)
    else:
        shutil.copy(source, dest_dir)

def covert_unc(host, path):
    """ Convert a file path on a host to a UNC path."""
    return ''.join(['\\\\', host, '\\', path.replace(':', '$')])
    
def wnet_connect(host, username, password):
    unc = ''.join(['\\\\', host])
    try:
        win32wnet.WNetAddConnection2(0, None, unc, None, username, password)
    except Exception as err:
        if isinstance(err, win32wnet.error):
            # Disconnect previous connections if detected, and reconnect.
            if err[0] == 1219:
                win32wnet.WNetCancelConnection2(unc, 0, 0)
                return wnet_connect(host, username, password)
        raise err

if __name__ == '__main__':
    for i in range(1, 17):
        netcopy('10.77.177.210', 'C:/performance_Logs - Copy/logs{}.csv'.format(i), \
            'E:/Users/Krishna.M/Downloads/Automation tool/raw logs from server', \
            username="StormTestUser", password="user_pwd", move=False)