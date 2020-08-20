import os
import time
import fnmatch
import logging
import threading
try:
    import paramiko
except:
    dbutils.library.installPyPI('paramiko')
    import paramiko


from concurrent.futures import ThreadPoolExecutor, ProcessPoolExecutor
logger = logging.getLogger(__name__)

class SFTPTransfer:
    """
    Simple wrapper class for SFTP transfer with multiprocessing Pool.
    :param hostname - Hostname of the source system
    :param user - login username
    :param passwd - login password

    :var ssh = ssh connection with the source server
    :var sftp = sftp connection through SSH
    """
    def __init__(self, hostname, user, passwd):
        self.host = hostname
        self.user = user
        self.passwd = passwd

        self.ssh = paramiko.SSHClient()
        self.ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())

        self.ssh.connect(hostname=self.host, username=self.user, password=self.passwd, auth_timeout=60, banner_timeout=60)
        self.sftp = self.ssh.open_sftp()

    def __enter__(self):
        return self

    def _file_check(self, absolute_path):
        try:
            _ = self.sftp.stat(absolute_path)
        except IOError:
            return False
        return True

    def _get_files_list(self, source, target, pattern, local=False):
        list_of_files = []
        if local:
            all_files = os.listdir(source)
        else:
            all_files = self.sftp.listdir(source)

        for src_file in all_files:
            if fnmatch.fnmatch(src_file, pattern):
                source_path = os.path.join(source, src_file)
                target_path = os.path.join(target, src_file)
                list_of_files.append([source_path, target_path])
        return list_of_files

    def __get_file(self, file_info, zero_byte=False):
        src_path, tgt_path = file_info

        stat = self._file_check(src_path)
        if stat:
            self.sftp.get(src_path,
                          tgt_path)
            return "SFTP successful : {}".format(file_info)
        elif zero_byte:
                open(tgt_path, 'a').close()
                return "SFTP successful : {}".format(file_info)
        else:
            raise FileNotFoundError("{} file does not exist".
                                    format(src_path))

    def __put_file(self, file_info, overwrite):
        src_path, tgt_path = file_info

        stat = os.path.exists(src_path)
        if stat:
            try:
                self.sftp.put(src_path, tgt_path)
                return "SFTP successful : {}".format(file_info)
            except PermissionError:
                if overwrite:
                    stat = self._file_check(tgt_path)
                    if stat:
                        self.sftp.remove(tgt_path)
                        self.sftp.put(src_path, tgt_path)
                        return "File {} uploaded successfully".format(src_path)
                    else:
                        raise PermissionError("Permission denied in target")
                else:
                    raise PermissionError("Target file exists or \
                        permission denied. Try again with overwrite=True")
        else:
            raise FileNotFoundError(
                "source file {} does not exist".format(local_file)
            )

    # Get list of files to be transferred and move through pool of processes
    # list_of_files - list of (source_path, target_path) combination
    def download_files(self, list_of_files=[],
                                source=None,
                                target=None,
                                file_pattern=None,
                                zero_byte=False):
        if not list_of_files:
            if source and target and file_pattern:
                list_of_files = self._get_files_list(source, target, file_pattern)
            else:
                raise ValueError("Either list_of_files or \
                  all of source, target, source_file_pattern must be passed")


        with ThreadPoolExecutor() as executor:
            for f in list_of_files:
              e = executor.submit(self.__get_file, f, zero_byte)
              logger.info(e.result())

    def upload_files(self, list_of_files=[],
                            source=None,
                            target=None,
                            file_pattern=None,
                            overwrite=False):

        if not list_of_files:
            if source and target and file_pattern:
                list_of_files = self._get_files_list(source, target, file_pattern, local=True)
            else:
                raise ValueError("Either list_of_files or \
                  all of source, target, source_file_pattern must be passed")

        with ThreadPoolExecutor() as executor:
            for f in list_of_files:
              e = executor.submit(self.__put_file, f, overwrite)
              logger.info(e.result())

    def close_conn(self):
        self.sftp.close()
        self.ssh.close()

    def __exit__(self, type, value, tb):
        self.close_conn()