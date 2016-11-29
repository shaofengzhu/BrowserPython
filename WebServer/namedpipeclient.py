import ctypes
import struct
import sys
import httphelper

class WinNamedPipeClient:
    PIPE_READMODE_MESSAGE = 0x00000002
    OPEN_EXISTING =         0x00000003
    GENERIC_READ =          0x80000000
    GENERIC_WRITE =         0x40000000
    ERROR_MORE_DATA =       234
    hk32 = ctypes.windll.LoadLibrary('kernel32.dll')

    @staticmethod
    def ctypes_handle(handle):
        if sys.maxsize > 2**32:
            return ctypes.c_ulonglong(handle)
        else:
            return ctypes.c_uint(handle)

    def __init__(self, name):
        self.name = name
        self.handle = WinNamedPipeClient.hk32['CreateFileA'](
            ctypes.c_char_p(b'\\\\.\\pipe\\' + bytes(name, 'utf8')),
            ctypes.c_uint(WinNamedPipeClient.GENERIC_READ | WinNamedPipeClient.GENERIC_WRITE),
            0,                      # no sharing
            0,                      # default security
            ctypes.c_uint(WinNamedPipeClient.OPEN_EXISTING),
            0,                      # default attributes
            0                       # no template file
        )

        if WinNamedPipeClient.hk32['GetLastError']() != 0:
            err = WinNamedPipeClient.hk32['GetLastError']()
            self.alive = False
            raise Exception('Pipe Open Failed [%s]' % err)
            return

        xmode = struct.pack('I', WinNamedPipeClient.PIPE_READMODE_MESSAGE)
        ret = WinNamedPipeClient.hk32['SetNamedPipeHandleState'](
            WinNamedPipeClient.ctypes_handle(self.handle),
            ctypes.c_char_p(xmode),
            ctypes.c_uint(0),
            ctypes.c_uint(0)
        )

        if ret == 0:
            err = WinNamedPipeClient.hk32['GetLastError']()
            self.alive = False
            raise Exception('Pipe Set Mode Failed [%s]' % err)
            return

        self.alive = True
        return

    def close(self):
        WinNamedPipeClient.hk32['CloseHandle'](WinNamedPipeClient.ctypes_handle(self.handle))
        alive = False

    def read(self) -> bytes:
        if not self.alive:
            raise Exception('Pipe Not Alive')
        message = bytes()
        buf = ctypes.create_string_buffer(4096)
        readMoreData = True
        while readMoreData:
            cnt = b'\x00\x00\x00\x00'
            ret = WinNamedPipeClient.hk32['ReadFile'](WinNamedPipeClient.ctypes_handle(self.handle), buf, 4096, ctypes.c_char_p(cnt), 0)
            readMoreData = False
            if ret == 0:
                err = WinNamedPipeClient.hk32['GetLastError']()
                if err == WinNamedPipeClient.ERROR_MORE_DATA:
                    readMoreData = True

            cnt = struct.unpack('I', cnt)[0]
            partialmessage = buf[0:cnt]
            message = message + partialmessage
        return message

    def write(self, rawmsg) -> None:
        if not self.alive:
            raise Exception('Pipe Not Alive')
        written = b'\x00\x00\x00\x00'
        ret = WinNamedPipeClient.hk32['WriteFile'](WinNamedPipeClient.ctypes_handle(self.handle), ctypes.c_char_p(rawmsg), ctypes.c_uint(len(rawmsg)), ctypes.c_char_p(written), ctypes.c_uint(0))
        if ret == 0:
            err = WinNamedPipeClient.hk32['GetLastError']()
            raise Exception('WriteFile Failed [%s]' % err)

if __name__ == "__main__":
    client = WinNamedPipeClient('fakeexcel')
    client.write(b'GET /activeWorkbook/sheets HTTP/1.0\n')
    reply = client.read()
    client.close()
    print(reply);


