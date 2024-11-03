#!/usr/bin/python
#
"""Provide drag and drop functionality between applications."""

import os
import sys

class Drop_Target_Interface(object):
    """Interface covering functions required for drag and drop functionality."""

    formats = {}

    def add_drop_format(self, format, handler):
        """Adds a format handler function."""
        if self.formats:
            self.formats[format] = handler
        else:
            self.formats = {format : handler}

    def get_handle(self):
        """Returns the handle used  for OS deag and drop registration."""
        print ("%s, should define get handle function." % (type(self)))
        return None
 
    def drag_enter(self, key, point, effect):
        """Handles drag enter event."""
        return winerror.S_OK
            
    def drag_over(self, key, point, effect):
        """Callback for drag over processing"""
        return winerror.S_OK

    def drag_leave(self):
        """Callback for drag leave processing"""
        return winerror.S_OK

    def drop(self, data, key, point, effect):
        """Callback for drop processing"""
        try:
            result = data.EnumFormatEtc(1).Next(100)
            for format in result:
                if format[0] in self.formats:
                    content = data.GetData(format)
                    if self.formats[format[0]](content, point):
                        break

                #if format[0] == win32con.CF_TEXT:
                #    text = data.GetData(format).data
                #    text = text.replace(r'\x00', '')
                #    self.drop_text(text[:-1], point)

                #if format[0] == win32con.CF_HDROP:
                #    result = data.GetData(format)
                #    files = []
                #    for file_number in range(win32api.DragQueryFile(result.data_handle)):
                #        files.append(win32api.DragQueryFile(result.data_handle, file_number))
                #    self.drop_file(files, point)
        except:
            print ("Unexpected error:", sys.exc_info()[0])
        return winerror.S_OK

if os.name == 'nt':
    import pythoncom
    import winerror
    import win32com.server.util
    import win32api
    import win32con
    import win32gui
    import win32clipboard

    TEXT = win32con.CF_TEXT
    FILES = win32con.CF_HDROP

    def get_files(result):
        """Returns a list of files from the drop results."""
        files = []
        for file_number in range(win32api.DragQueryFile(result.data_handle)):
            files.append(win32api.DragQueryFile(result.data_handle, file_number))
        return files

    standard_formats = [win32con.CF_BITMAP,
                        win32con.CF_DIB,
                        win32con.CF_DIF,
                        win32con.CF_DSPBITMAP,
                        win32con.CF_DSPENHMETAFILE,
                        win32con.CF_DSPMETAFILEPICT,
                        win32con.CF_DSPTEXT,
                        win32con.CF_ENHMETAFILE,
                        win32con.CF_GDIOBJFIRST,
                        win32con.CF_HDROP,
                        win32con.CF_LOCALE,
                        win32con.CF_METAFILEPICT,
                        win32con.CF_OEMTEXT,
                        win32con.CF_OWNERDISPLAY,
                        win32con.CF_PALETTE,
                        win32con.CF_PENDATA,
                        win32con.CF_PRIVATEFIRST,
                        win32con.CF_RIFF,
                        win32con.CF_SYLK,
                        win32con.CF_TEXT,
                        win32con.CF_WAVE,
                        win32con.CF_TIFF,
                        win32con.CF_UNICODETEXT]

    class Drop_Target:
        """Wrapper for the IDropTarget functionality"""
        _public_methods_ = [ 'DragEnter', 'DragOver', 'DragLeave', 'Drop']
        _com_interfaces_ = [ pythoncom.IID_IDropTarget ]
        _reg_progid_ = "Python.planner"
        _reg_clsid_ = pythoncom.CreateGuid()
        _ole_initialised = False
        _registered = []
        __shared_state = {}

        def __init__(self, widget):
            """Set up the COM elements of required to handle drag and drop from another
            application. Assumes the widget has an associated kernel window to receive events.
            
            """
            self.helper = None
            self.pending = None

            handle = widget.get_handle()
            if handle:
                com_object = win32com.server.util.wrap(self) #, useDispatcher = 1
                if not Drop_Target._ole_initialised:
                    pythoncom.OleInitialize()
                    Drop_Target._ole_initialised = True
                self.target = widget
                handle = widget.get_handle()
                if handle not in Drop_Target._registered:
                    pythoncom.RegisterDragDrop(handle, com_object)
                    win32gui.DragAcceptFiles(handle, True)
                    Drop_Target._registered.append(handle)

        def remove(self, window):
            """Remove this window as a drop target."""
            handle = window.get_handle()
            if handle in Drop_Target._registered :
                win32gui.DragAcceptFiles(handle, False)
                Drop_Target._registered.remove(handle)
                if self.target == window:
                    self.target = None

        def DragEnter(self, data, key, point, effect):
            """Callback for drag enter processing"""
            try:
                if self.helper:
                    self.helper.DragEnter(self.window.get_handle(), data, point, effect)
                effect = 4
                self.target.drag_enter(key, point, effect)
            except:
                print ("Unexpected error:", sys.exc_info()[0])
            finally:
                return winerror.S_OK

        def DragOver(self, key, point, effect):
            """Callback for drag over processing"""
            try:
                if self.helper:
                    self.helper.DragOver(point, effect)
                effect = 4
                self.target.drag_over(key, point, effect)
            except:
                print ("Unexpected error:", sys.exc_info())
            finally:
                return winerror.S_OK

        def DragLeave(self):
            """Callback for drag leave processing"""
            try:
                if self.helper:
                    self.helper.DragLeave()
                self.target.drag_leave_()
            except:
                print ("Unexpected error:", sys.exc_info()[0])
            finally:
                return winerror.S_OK

        def Drop(self, data, key, point, effect):
            """Callback for drop processing"""
            try:
                if self.target:
                    self.target.drop(data, key, point, effect)
            except:
                print ("Unexpected error:", sys.exc_info()[0])
            return winerror.S_OK

else:
    import gtk

    TEXT = 80
    FILES = None

    def get_files(result):
        """Returns a list of files from the drop results."""
        files = []
        return files

    class Drop_Target:

        def null_function(self, *args):
            """Do nothing..."""
            pass

        def __init__(self, window):
            """Set up the COM elements of required to handle drag and drop from another
            application"""
            self.helper = None

            self.window = window
            TARGET_TYPE_TEXT = 80
            targets = [( "text/plain", 0, TARGET_TYPE_TEXT )] 
            self.window.drag_dest_set(gtk.DEST_DEFAULT_ALL, targets, gtk.gdk.ACTION_COPY |  gtk.gdk.ACTION_LINK)
            self.window.connect("drag_data_received", self.window.drop_text)

        def DragEnter(self, data, key, point, effect):
            """Callback for drag enter processing"""
            try:
                if self.helper:
                    self.helper.DragEnter(self.window.get_handle(), data, point, effect)
                effect = 4
                self.window.drag_enter(key, point, effect)
            except:
                print ("Unexpected error:", sys.exc_info()[0])
            finally:
                return True

        def DragOver(self, key, point, effect):
            """Callback for drag over processing"""
            try:
                if self.helper:
                    self.helper.DragOver(point, effect)
                effect = 4
                self.window.drag_over(key, point, effect)
            except:
                print ("Unexpected error:", sys.exc_info()[0])
            finally:
                return True

        def DragLeave(self):
            """Callback for drag leave processing"""
            try:
                if self.helper:
                    self.helper.DragLeave()
                self.window.drag_leave()
            except:
                print ("Unexpected error:", sys.exc_info()[0])
            finally:
                return True

        def Drop(self, data, key, point, effect):
            """Callback for drop processing"""
            try:
                self.window.drop()
            except:
                print ("Unexpected error:", sys.exc_info()[0])
            return True

class Test_Module:
    """Test code run when this module is executed directly."""

    def __init__ (self):
        """Need to develop test case!!!"""
        pass

# if this file is executed run some simple tests
if __name__ == "__main__":
    Test_Module()
