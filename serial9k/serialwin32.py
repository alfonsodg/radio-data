#-*- coding: utf-8 -*-

import time
import ctypes
import array
import struct
import copy
import winioctlcon as win32con
import win32api
import win32file
import win32event

from serial.serialutil import SerialException, FileLike, portNotOpenError, \
        writeTimeoutError, LF
from serial.serialwin32 import Win32Serial

SPOOL_BUFFER = 1650 # bytes

FILE_DEVICE_SERIAL_PORT = 0x0000001B
METHOD_BUFFERED = 0
FILE_ANY_ACCESS = 0

def ctl_code(device_type, function, method, access):
    return device_type << 16 | access << 14 | function << 2 | method

# defining some raw IOCTL constants for serial-device control requests
IOCTL_SERIAL_GET_BAUD_RATE = getattr(win32con, 
    'IOCTL_SERIAL_GET_BAUD_RATE', ctl_code(FILE_DEVICE_SERIAL_PORT, 20,
        METHOD_BUFFERED, FILE_ANY_ACCESS))
IOCTL_SERIAL_GET_LINE_CONTROL = getattr(win32con, 
    'IOCTL_SERIAL_GET_LINE_CONTROL',ctl_code(FILE_DEVICE_SERIAL_PORT, 21,
        METHOD_BUFFERED, FILE_ANY_ACCESS))
IOCTL_SERIAL_GET_CHARS = getattr(win32con, 
    'IOCTL_SERIAL_GET_CHARS', ctl_code(FILE_DEVICE_SERIAL_PORT, 22,
        METHOD_BUFFERED, FILE_ANY_ACCESS))
IOCTL_SERIAL_GET_HANDFLOW = getattr(win32con, 
    'IOCTL_SERIAL_GET_HANDFLOW', ctl_code(FILE_DEVICE_SERIAL_PORT, 24,
        METHOD_BUFFERED, FILE_ANY_ACCESS))
IOCTL_SERIAL_GET_PROPERTIES = getattr(win32con, 
    'IOCTL_SERIAL_GET_PROPERTIES', ctl_code(FILE_DEVICE_SERIAL_PORT, 29,
        METHOD_BUFFERED, FILE_ANY_ACCESS))

IOCTL_SERIAL_CLR_RTS = getattr(win32con, 
    'IOCTL_SERIAL_CLR_RTS', ctl_code(FILE_DEVICE_SERIAL_PORT, 13,
        METHOD_BUFFERED, FILE_ANY_ACCESS))
IOCTL_SERIAL_CLR_DTR = getattr(win32con, 
    'IOCTL_SERIAL_CLR_DTR', ctl_code(FILE_DEVICE_SERIAL_PORT, 10,
        METHOD_BUFFERED, FILE_ANY_ACCESS))

IOCTL_SERIAL_SET_BAUD_RATE = getattr(win32con, 
    'IOCTL_SERIAL_SET_BAUD_RATE', ctl_code(FILE_DEVICE_SERIAL_PORT, 1,
        METHOD_BUFFERED, FILE_ANY_ACCESS))
IOCTL_SERIAL_SET_LINE_CONTROL = getattr(win32con, 
    'IOCTL_SERIAL_SET_LINE_CONTROL', ctl_code(FILE_DEVICE_SERIAL_PORT, 3,
        METHOD_BUFFERED, FILE_ANY_ACCESS))
IOCTL_SERIAL_SET_CHARS = getattr(win32con,
    'IOCTL_SERIAL_SET_CHARS', ctl_code(FILE_DEVICE_SERIAL_PORT, 23,
        METHOD_BUFFERED, FILE_ANY_ACCESS))
IOCTL_SERIAL_SET_HANDFLOW = getattr(win32con, 
    'IOCTL_SERIAL_SET_HANDFLOW', ctl_code(FILE_DEVICE_SERIAL_PORT, 25,
        METHOD_BUFFERED, FILE_ANY_ACCESS))
IOCTL_SERIAL_SET_TIMEOUTS = getattr(win32con, 
    'IOCTL_SERIAL_SET_TIMEOUTS', ctl_code(FILE_DEVICE_SERIAL_PORT, 7,
        METHOD_BUFFERED, FILE_ANY_ACCESS))
IOCTL_SERIAL_SET_QUEUE_SIZE = getattr(win32con, 
    'IOCTL_SERIAL_SET_QUEUE_SIZE', ctl_code(FILE_DEVICE_SERIAL_PORT, 2,
        METHOD_BUFFERED, FILE_ANY_ACCESS))
IOCTL_SERIAL_SET_RTS = getattr(win32con, 
    'IOCTL_SERIAL_SET_RTS', ctl_code(FILE_DEVICE_SERIAL_PORT, 12,
        METHOD_BUFFERED, FILE_ANY_ACCESS))
IOCTL_SERIAL_PURGE = getattr(win32con, 
    'IOCTL_SERIAL_PURGE', ctl_code(FILE_DEVICE_SERIAL_PORT, 19,
        METHOD_BUFFERED, FILE_ANY_ACCESS))
IRP_MJ_WRITE = getattr(win32con, 'IRP_MJ_WRITE', 0x04)

# defining structures
#   typedef struct _SERIAL_BAUD_RATE {
#       ULONG BaudRate;
#   };
SERIAL_BAUD_RATE_STRUCT = array.array('B', [0]*struct.calcsize('L'))

#   typedef struct _SERIAL_LINE_CONTROL {
#       UCHAR StopBits;
#       UCHAR Parity;
#       UCHAR WordLength;
#   };
SERIAL_LINE_CONTROL_STRUCT = array.array('B', [0]*struct.calcsize('BBB'))

#   typedef struct _SERIAL_CHARS {
#       UCHAR EofChar;
#       UCHAR ErrorChar;
#       UCHAR BreakChar;
#       UCHAR EventChar;
#       UCHAR XonChar;
#       UCHAR XoffChar;
#   };
SERIAL_CHARS_STRUCT = array.array('B', [0]*struct.calcsize('BBBBBB'))

#   typedef struct _SERIAL_HANDFLOW {
#       ULONG ControlHandShake;
#       ULONG FlowReplace;
#       LONG XonLimit;
#       LONG XoffLimit;
#   };
SERIAL_HANDFLOW_STRUCT = array.array('B', [0]*struct.calcsize('LLll'))

#   typedef struct _SERIAL_TIMEOUTS {
#       ULONG ReadIntervalTimeout;
#       ULONG ReadTotalTimeoutMultiplier;
#       ULONG ReadTotalTimeoutConstant;
#       ULONG WriteTotalTimeoutMultiplier;
#       ULONG WriteTotalTimeoutConstant;
#   };
SERIAL_TIMEOUTS_STRUCT = array.array('B', [0]*struct.calcsize('LLLLL'))

#   typedef struct _SERIAL_QUEUE_SIZE {
#       ULONG InSize;
#       ULONG OutSize;
#   };
SERIAL_QUEUE_SIZE_STRUCT = array.array('B', [0]*struct.calcsize('LL'))

#   typedef struct _SERIAL_COMMPROP {
#       USHORT PacketLength;
#       USHORT PacketVersion;
#       ULONG ServiceMask;
#       ULONG Reserved1;
#       ULONG MaxTxQueue;
#       ULONG MaxRxQueue;
#       ULONG MaxBaud;
#       ULONG ProvSubType;
#       ULONG ProvCapabilities;
#       ULONG SettableParams;
#       ULONG SettableBaud;
#       USHORT SettableData;
#       USHORT SettableStopParity;
#       ULONG CurrentTxQueue;
#       ULONG CurrentRxQueue;
#       ULONG ProvSpec1;
#       ULONG ProvSpec2;
#       WCHAR ProvChar[1];
#   };
SERIAL_COMMPROP_STRUCT = array.array('B',
        [0]*struct.calcsize('IILLLLLLLLLIILLLLI'))

# ULONG
SERIAL_PURGE_STRUCT = array.array('B', [0]*struct.calcsize('L'))

class Serial9kWin32(Win32Serial):
    """
    Serial port implementation for Win32 based on ctypes.
    """
    def open(self):
        """
        Open port with current settings. This may throw a SerialException if the
        port cannot be opened.
        """
        if self._port is None:
            raise SerialException('Port must be configured before it can be ' \
                'used.')
        # the "\\.\COMx" format is required for devices other than COM1-COM8
        # not all versions of windows seem to support this properly
        # so that the first few ports are used with the DOS device name
        port = self.portstr
        try:
            if port.upper().startswith('COM') and int(port[3:]) > 8:
                port = '\\\\.\\' + port
        except ValueError:
            # for like COMnotanumber
            pass

        self._rtsState = win32file.RTS_CONTROL_ENABLE
        self._dtrState = win32file.DTR_CONTROL_ENABLE

        self.hComPort = win32file.CreateFile(port,
               win32file.GENERIC_READ | win32file.GENERIC_WRITE,
               0, # exclusive access
               None, # no security
               win32file.OPEN_EXISTING,
               win32file.FILE_ATTRIBUTE_NORMAL | win32file.FILE_FLAG_OVERLAPPED,
               0)

        if self.hComPort == win32file.INVALID_HANDLE_VALUE:
            self.hComPort = None    # 'cause __del__ is called anyway
            raise SerialException('could not open port %s: %s' % (self.portstr, 
                ctypes.WinError()))

        self._reconfigurePort()

        win32file.CloseHandle(self.hComPort)

        self.hComPort = win32file.CreateFile(port,
               win32file.GENERIC_READ | win32file.GENERIC_WRITE,
               0, # exclusive access
               None, # no security
               win32file.OPEN_EXISTING,
               win32file.FILE_ATTRIBUTE_NORMAL | win32file.FILE_FLAG_OVERLAPPED,
               0)

        if self.hComPort == win32file.INVALID_HANDLE_VALUE:
            self.hComPort = None    # 'cause __del__ is called anyway
            raise SerialException('could not open port %s: %s' % (self.portstr, 
                ctypes.WinError()))

        self._reconfigurePort()
        self._purge()

        self._overlappedRead = win32file.OVERLAPPED()
        self._overlappedRead.hEvent = win32event.CreateEvent(None, 1, 0, None)
        self._overlappedWrite = win32file.OVERLAPPED()
        self._overlappedWrite.hEvent = win32event.CreateEvent(None, 0, 0, None)
        self._isOpen = True

    def _reconfigurePort(self):
        """
        Set communication parameters on opened port.
        """
        if not self.hComPort:
            raise SerialException('Can only operate on a valid port handle')

        self._first_payload()
        data = self._first_payload()
        self._second_payload(data)
        self._third_payload()
        self._first_payload()
        data = self._first_payload()
        self._second_payload(data)

    def _first_payload(self):
        data = copy.deepcopy(SERIAL_BAUD_RATE_STRUCT)
        win32file.DeviceIoControl(self.hComPort, IOCTL_SERIAL_GET_BAUD_RATE, 
                None, data, None)
        data = copy.deepcopy(SERIAL_LINE_CONTROL_STRUCT)
        win32file.DeviceIoControl(self.hComPort, IOCTL_SERIAL_GET_LINE_CONTROL,
                None, data, None)
        data = copy.deepcopy(SERIAL_CHARS_STRUCT)
        win32file.DeviceIoControl(self.hComPort, IOCTL_SERIAL_GET_CHARS,
                None, data, None)
        data = copy.deepcopy(SERIAL_HANDFLOW_STRUCT)
        return win32file.DeviceIoControl(self.hComPort, 
                IOCTL_SERIAL_GET_HANDFLOW,
                None, data, None)

    def _second_payload(self, handshake):
        # 115200: 00 C2 01 00
        # 9600: 80 25 00 00
        # 1200: B0 04 00 00
        # TODO: We are assuming 9600 all times
        data = copy.deepcopy(SERIAL_BAUD_RATE_STRUCT)
        if self.baudrate == 1200:
            data[0] = 0xB0
            data[1] = 0x04
            data[2] = 0x00
            data[3] = 0x00
        elif self.baudrate == 9600:
            data[0] = 0x80
            data[1] = 0x25
            data[2] = 0x00
            data[3] = 0x00
        else:
            data[0] = 0x00
            data[1] = 0xC2
            data[2] = 0x01
            data[3] = 0x00
        win32file.DeviceIoControl(self.hComPort, IOCTL_SERIAL_SET_BAUD_RATE, 
                data, 0, None)
        win32file.DeviceIoControl(self.hComPort, IOCTL_SERIAL_CLR_RTS,
                None, 0, None)
        win32file.DeviceIoControl(self.hComPort, IOCTL_SERIAL_CLR_DTR,
                None, 0, None)

        # 115200: 02 00 08
        # 9600: 00 00 08
        # 1200: 00 02 08
        # TODO: We are assuming 9600 all times
        data = copy.deepcopy(SERIAL_LINE_CONTROL_STRUCT)
        if self.baudrate == 1200:
            data[0] = 0x00
            data[1] = 0x02
            data[2] = 0x08
        elif self.baudrate == 9600:
            data[0] = 0x00
            data[1] = 0x00
            data[2] = 0x08
        else:
            data[0] = 0x02
            data[1] = 0x00
            data[2] = 0x08
        win32file.DeviceIoControl(self.hComPort, IOCTL_SERIAL_SET_LINE_CONTROL,
                data, 0, None)

        # 00 00 00 00 11 13
        data = copy.deepcopy(SERIAL_CHARS_STRUCT)
        data[0] = 0x00
        data[1] = 0x00
        data[2] = 0x00
        data[3] = 0x00
        data[4] = 0x11
        data[5] = 0x13
        win32file.DeviceIoControl(self.hComPort, IOCTL_SERIAL_SET_CHARS,
                data, 0, None)

        # 00 00 00 00 00 00 00 00 C0 86 00 00 B0 21 00 00
        #   data = copy.deepcopy(SERIAL_HANDFLOW_STRUCT)
        #   data[0] = 0x00
        #   data[1] = 0x00
        #   data[2] = 0x00
        #   data[3] = 0x00
        #   data[4] = 0x00
        #   data[5] = 0x00
        #   data[6] = 0x00
        #   data[7] = 0x00
        #   data[8] = 0xC0
        #   data[9] = 0x86
        #   data[10] = 0x00
        #   data[11] = 0x00
        #   data[12] = 0xB0
        #   data[13] = 0x21
        #   data[14] = 0x00
        #   data[15] = 0x00
        win32file.DeviceIoControl(self.hComPort, IOCTL_SERIAL_SET_HANDFLOW,
                handshake, 0, None)

    def _third_payload(self):
        # 1E 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00
        data = copy.deepcopy(SERIAL_TIMEOUTS_STRUCT)
        data[0] = 0x1E
        data[1] = 0x00
        data[2] = 0x00
        data[3] = 0x00
        data[4] = 0x00
        data[5] = 0x00
        data[6] = 0x00
        data[7] = 0x00
        data[8] = 0x00
        data[9] = 0x00
        data[10] = 0x00
        data[11] = 0x00
        data[12] = 0x00
        data[13] = 0x00
        data[14] = 0x00
        data[15] = 0x00
        data[16] = 0x00
        data[17] = 0x00
        data[18] = 0x00
        data[19] = 0x00
        win32file.DeviceIoControl(self.hComPort, IOCTL_SERIAL_SET_TIMEOUTS,
                data, 0, None)

        # 30 75 00 00 C0 C6 2D 00
        data = copy.deepcopy(SERIAL_QUEUE_SIZE_STRUCT)
        data[0] = 0x30
        data[0] = 0x75
        data[0] = 0x00
        data[0] = 0x00
        data[0] = 0xC0
        data[0] = 0xC6
        data[0] = 0x2D
        data[0] = 0x00
        win32file.DeviceIoControl(self.hComPort, IOCTL_SERIAL_SET_QUEUE_SIZE,
                data, 0, None)

        data = copy.deepcopy(SERIAL_COMMPROP_STRUCT)
        win32file.DeviceIoControl(self.hComPort, 
                IOCTL_SERIAL_GET_PROPERTIES,
                None, data, None)

    def _purge(self):
        # 0A 00 00 00
        data = copy.deepcopy(SERIAL_PURGE_STRUCT)
        data[0] = 0x0A
        data[1] = 0x00
        data[2] = 0x00
        data[3] = 0x00
        win32file.DeviceIoControl(self.hComPort, IOCTL_SERIAL_PURGE, data, 0, 
                None)

    def close(self):
        """Close port"""
        if self._isOpen:
            if self.hComPort:
                # Close COM-Port:
                win32file.CloseHandle(self.hComPort)
                win32file.CloseHandle(self._overlappedRead.hEvent)
                win32file.CloseHandle(self._overlappedWrite.hEvent)
                self.hComPort = None
            self._isOpen = False

    def write(self, data):
        """
        Output the given string over the serial port.
        """
        if not self.hComPort: raise portNotOpenError
        data = bytes(data)
        if data:
            self._purge()
            win32file.DeviceIoControl(self.hComPort, IOCTL_SERIAL_SET_RTS,
                    None, 0, None)
            time.sleep(1) # 8000 ms timeout
            start = 0
            end = SPOOL_BUFFER
            chunk = data[start:end]
            # header
            written, err = win32file.WriteFile(self.hComPort, 
                    struct.pack('<Q', len(data)), self._overlappedWrite)
            err = win32file.GetOverlappedResult(self.hComPort, 
                    self._overlappedWrite, True)
            data += struct.pack('<I', 0xFFFFFFFF) + '\n' * 1650 * 2
            while chunk:
                written, err = win32file.WriteFile(self.hComPort, chunk, 
                        self._overlappedWrite)
                err = win32file.GetOverlappedResult(self.hComPort, 
                        self._overlappedWrite, True)
                start = end
                end += SPOOL_BUFFER
                chunk = data[start:end]

            return written
        else:
            return 0

    def read(self, size=SPOOL_BUFFER):
        """
        Read size bytes from the serial port. If a timeout is set it may
        return less characters as requested. With no timeout it will block
        until the requested number of bytes is read.
        """
        if not self.hComPort: raise portNotOpenError
        win32event.ResetEvent(self._overlappedRead.hEvent)

        data = ''
        has_header = False
        while True:
            err, chunk = win32file.ReadFile(self.hComPort, SPOOL_BUFFER, 
                    self._overlappedRead)
            err = win32event.WaitForSingleObject(self._overlappedRead.hEvent, 
                    win32event.INFINITE)
            data += str(chunk)
            if not has_header:
                data_size = struct.unpack('<Q', data[:8])[0]
                has_header = True
            chunk = data[8:8+data_size]
            if has_header and len(chunk) == data_size:
                break
        return chunk

try:
    import io
except ImportError:
    # classic version with our own file-like emulation
    class Serial9k(Serial9kWin32, FileLike):
        pass
else:
    # io library present
    class Serial9k(Serial9kWin32, io.RawIOBase):
        pass
