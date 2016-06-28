import os
import os.path
import shutil
import sys
import win32wnet

import argparse

parser = argparse.ArgumentParser(description='Grab log files from production server and store it under logs folder')
parser.add_argument('Server_IP', nargs='+', help='The IP address for the server storing log files')
parser.add_argument('User_name', nargs='+', help='The user name for the storage server')
parser.add_argument('Password', nargs='+', help='The Password name for the storage server')
parser.add_argument('remote_path_to_log_files', nargs='+', help='The path on the remote machine that stores the log files')

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
    try:
        server_IP = sys.argv[1]
        server_UName = sys.argv[2]
        server_password = sys.argv[3]
        remote_path_to_log_files = sys.argv[4]
        for i in range(1, 17):
            netcopy(server_IP, '{}/logs{}.csv'.format(remote_path_to_log_files , i), \
                'logs', username=server_UName, password=server_password, move=False)
        print("\n\nCopy Successful!!!\n")
    except IndexError:
        parser.print_help()