#!/usr/bin/python
#
# Simple time tracking application with system tray icon
#

# Options
#
defaults = {'data_dir'          : '~/task_tracker/',
            'config_file'       : 'configuration.xml',
            'task_file'         : 'tasks.xml',
            'session_file'      : 'session.xml',
            'daily_backup'      : True,
            'startup_task'      : 'previous',
            'sort_tasks'        : False,
            'hierarchical_list' : True,
            'submenus_on_top'   : True,
            'import_tasks'      : False,
            'pause_on_exit'     : True,
            'show_timesheet'    : False,
            'inactive_nag'      : True,
            'priority_expanded' : True,
            'recent_expanded'   : False,
            'len_recent_tasks'  : 5,
            'include_future'    : False,
            'booking_tips_full' : False,
            'outlook_import'    : True,
            'timesheet' : {'show' : True},
            'notepad' : {'show' : False},
            'todo' : {'show'         : True,
                      'full_name'    : False,
                      'scheduled'    : False,
                      'opacity'      : 50,
                      'keep_above'   : True,
                      'position'     : 1,
                      'x'            : 0,
                      'y'            : 0,
                      'width'        : 0,
                      'height'       : 0,
                      'timeout'      : 5000,
                      'size'         : 5,
                      'sort_key'     : 'priority',
                      'current_first': True,
                      'show_buttons' : True},
            'stay_on_top'       : True,
            'show_unbilled'     : False,
            'shorten_names'     : True,
            'verbosity'         : 11,
            'day_start'         : '08:00',
            'day_end'           : '18:00',
            'hours_per_week'    : 38,
            'days_per_week'     : 5.0,
            'inactive_is_complete' : True,
            'pause_for_lunch'   : True,
            'lunch_duration'    : 1.0,
            'lunch_start'       : '12:00',
            'lunch_end'         : '13:00',
            'date_format'       : '%d %B %Y',
            'time_format'       : '%H:%M:%S',
            'merge_event_window': 5.0,
            'allow_unfinished'  : True,
            'join_activities'   : True,
            'allow_overnight'   : False,
            'remove_null_events': True,
            'division_spacing'  : 50.0,
            'image_root'        : 'oxygen/',
            'image_small'       : '16x16/',
            'image_medium'      : '22x22/'}

# Required modules
#
import datetime

import operator
import glob
import pygtk
pygtk.require('2.0')
import gobject
import gtk
import cairo
import gzip
import hashlib
import math
import os
import pango
import re
import shutil
import socket
import subprocess
import sys
import urllib
import urlparse
import weakref
import xml.dom.minidom
import dragndrop

from optparse import OptionParser

if os.name == 'nt':
    import pywintypes
    import win32com.client
    import win32api
    import win32security
elif os.name == 'posix':
    import pwd

weekdays = [0, 1, 2, 3, 4]

class User(object):
    """Manager of user indentities."""

    @classmethod
    def get_id(cls):
        """Return a string representing a user ID."""
        if os.name == 'nt':
            name = win32api.GetUserName()
            sid = win32security.LookupAccountName(None, name)
            return (name, sid)
        elif os.name == 'posix':
            uid = os.getuid()
            name = pwd.getpwuid(uid)[0]
            return (name, uid)
        else:
            print ("Unknown OS!")

def find(func, seq):
    """Return first item in sequence where func(item) == True."""
    for item in seq:
        if func(item):
            return item

def get_decimal_hours(duration):
    """Return the duration of as decimal hours."""
    return duration.days * 24.0 + duration.seconds / 3600.0

def try_conversion(patterns, string):
    """Attempt to convert the string to a date / time using the supplied formats."""
    match = None
    for pattern in patterns:
        try:
            match = datetime.datetime.strptime(string, pattern)
            break
        except ValueError:
            pass

    return match

def string_to_date(string, broken_order = False):
    """Returns a datetime.date object representing the date in the string."""

    date = try_conversion(['%c', '%Y-%m-%dT%H:%M:%S', '%d %b %y', '%d %B %y', '%d %b %Y', '%d %B %Y', '%B %d %Y', '%B %d %Y', '%b %d %Y', '%b %d %Y'], string)

    if not date:
        pattern1 = re.compile(r"(?P<day>[0-9]{1,2})[^0-9]+(?P<month>[0-9]{1,2})[^0-9]+(?P<year>([0-9]{1,2}){1,2})")
        pattern2 = re.compile(r"(?P<year>([0-9]{1,2}){1,2})[^0-9]+(?P<month>[0-9]{1,2})[^0-9]+(?P<day>[0-9]{1,2})")
        match = pattern1.match(string)
        if not match:
            match = pattern2.match(string)
        if match:
            year = int(match.group('year'))
            if year < 100:
               year += 2000
            month = int(match.group('month'))
            day = int(match.group('day'))
            if broken_order:
                day, month = month, day
            if month > 12 and day < 12:
                month, day = day, month
            date = datetime.datetime(year = year, month = month, day = day)

    if date:
        date = datetime.datetime.combine(date, datetime.time())

    return date

def string_to_time(string):
    """convert a string containing an iso format time to a datetime object."""

    # Ignore microseconds
    if '.' in string:
        string = string[:string.find('.')]
    time = try_conversion(['%c', '%Y-%m-%dT%H:%M:%S'], string)

    if not time:
        try:
            date = string_to_date(string)
            time = try_conversion(['%c', '%d %b %y', '%d %B %y', '%d %b %Y', '%d %B %Y'], string)
            time = datetime.datetime(date.year, date.month, date.day,
                                     int(string[11:13]), int(string[14:16]), int(string[17:19]), int(string[20:]))
        except:
            if string[20:]:
                time = datetime.datetime(int(string[0:4]), int(string[5:7]), int(string[8:10]),
                                         int(string[11:13]), int(string[14:16]), int(string[17:19]), int(string[20:]))
            else:
                time = datetime.datetime(int(string[0:4]), int(string[5:7]), int(string[8:10]),
                                         int(string[11:13]), int(string[14:16]), int(string[17:19]))
    return time


def string_to_timedelta(string):
    """convert a string to a time delta."""
    return datetime.timedelta(hours = float(string))

def timedelta_to_string(delta):
    """Returns a string representation of the timedelta."""
    result = []
    days = int(delta.days)
    if days:
        result.append('%d days' % (days))

    hours = int(delta.seconds / 3600)
    if hours:
        result.append('%d hr' % (hours))

    minutes = int((delta.seconds - hours * 3600) / 60)
    if minutes:
        result.append('%d min' % (minutes))

    seconds = int(delta.seconds - hours * 3600 - minutes * 60)
    if seconds:
        result.append('%d sec' % (seconds))
    return ' '.join(result)

def build_menu(menu_definition, bar = False):
    """Create a popup menu for the task elements of the tree.
    Menu contains, add task, add sub-task, delete task.

    """
    if bar:
        menu = gtk.MenuBar()
    else:
        menu = gtk.Menu()

    group = None

    for definition in menu_definition:
        text = definition[0]
        function = definition[1]
        if len(definition) > 2:
            data = definition[2]
        else:
            data = None
        if text == '---':
            item = gtk.SeparatorMenuItem()
            group = None
        elif text.startswith('>>>'):
            group = None
            item = gtk.MenuItem(text[3:])
            item.set_submenu(build_menu(function))
        elif text.startswith('oo'):
            item = gtk.RadioMenuItem(group, text[3:])
            if text[2] == 'x':
                item.set_active(True)
            item.connect("activate", function, data)
            group = item
        elif text.startswith('ox'):
            item = gtk.CheckMenuItem(text[3:])
            if text[2] == 'x':
                item.set_active(True)
            if data != None:
                item.connect("activate", function, data)
            else:
                item.connect("activate", function)
            group = None
        else:
            group = None
            item = gtk.MenuItem(text)
            if data != None:
                item.connect("activate", function, data)
            else:
                item.connect("activate", function)
        menu.append(item)
        item.show()
    return menu

def expand_list(base_list):
    """Returns a list of the given items with all decendents included."""
    item_list = []
    for item in base_list:
        item_list.append(item)
        item_list.extend(expand_list(item.children))
    return item_list

def accumulate_hours(task_manager):
    """Adds up the hours in all the session files and assigns them to the corresponding task."""

    session = {}
    path, extension = os.path.splitext(config.get('data_dir') + config.get('session_file'))
    filename = os.path.expanduser("%s_*%s" % (path, extension))

    for session_file in glob.glob(filename):
        session[session_file] = Activity_Manager(task_manager, None, filename = session_file)
        for activity in session[session_file].activities:
            activity.task.allocated_time += activity.get_duration()

    task_list = expand_list(task_manager.tasks)

    for task in task_list:
        print (task.name, get_decimal_hours(task.allocated_time))

class Configuration(dict):
    """Manage the application configuration."""

    def __init__(self, options):
        """Set the command line configuration options."""

        if options.data_dir:
           defaults['data_dir'] = options.data_dir + '/'

        data_dir = os.path.normpath(defaults['data_dir']) + '/'

        if options.config_file:
            if options.data_dir:
               config_file = defaults['data_dir'] + options.config_file
            else:
               config_file = options.config_file
        else:
           config_file = data_dir + defaults['config_file']
        config_file = os.path.normpath(config_file)

        # Create a copy of the default items
        for key, value in defaults.items():
            self[key] = value
        self.load(config_file)
        
        # Overwrite settings from the command line
        if options.data_dir:
            self['data_dir'] = data_dir

        if options.image_root:
            self['image_root'] = options.image_root + '/'

        if not os.path.exists(self['image_root']):
            if os.path.exists('../' + self['image_root']):
                self['image_root'] = '../' + self['image_root']
            else:
                # Cannot display GUI error before completing configuration
                # initialisation, so print (to console)
                print ('Incorrect Path', \)
                      "Image path %s doesn't exist" % (self['image_root'])

        self['image_path_16'] = self['image_root'] + '/' + self['image_small']
        self['image_path_16'] = os.path.normpath(self['image_path_16']) + '/actions/'
        self['image_path_22'] = self['image_root'] + '/' + self['image_medium']
        self['image_path_22'] = os.path.normpath(self['image_path_22']) + '/actions/'

        if options.initial_session:
            self['initial_session'] = options.initial_session

        if options.show_timesheet:
            self['timesheet']['show'] = True

        if options.session_file:
            self['custom_session_file'] = options.session_file

        if options.verbose:
            for key in self.keys():
                print ("configuration[%s] = %s" % (key, self[key]))
        self.__dict__['initialised'] = True

    def load(self, config_file):
        """Updates the settings based on the contents of the supplied xml file."""
        config_file = os.path.expanduser(config_file)
        if os.path.exists(config_file):
            pattern = re.compile("\s*(<|>)\s*")
            file = open(config_file)
            xml_string = file.read()
            file.close()
            # Remove pretty printing
            xml_string = pattern.sub(r'\1', xml_string)
            document = xml.dom.minidom.parseString(xml_string)
            for key, value in self.decode(document.firstChild).items():
                self[key] = value

    def decode(self, parent):
        """Return a dictionary containing the decoded xml node data."""
        results = {}
        for node in parent.childNodes:
            value = str(node.firstChild.nodeValue).strip()
            if node.hasAttribute('type'):
                node_type = node.getAttribute('type')
            else:
                node_type = None
                
            if node_type == 'parent' or node.firstChild.hasChildNodes():
                results[node.tagName] = self.decode(node)
            elif node_type == 'bool' or (not node_type and
                                         value.lower() in [u'false', u'no', u'true', u'yes']):
                results[node.tagName] = False if value.lower() in [u'false', u'no'] else True
            elif node_type == 'int' or (not node_type and value.isdigit()):
                results[node.tagName] = int(value)
            elif node_type == 'float' or (not node_type and
                                           value.replace('.', '', 1).isdigit()):
                results[node.tagName] = float(value)
            else:
                results[node.tagName] = value
        return results

    def encode(self, data, name, document):
        """Convert the contents of the data into an xml node."""
        parent_node = document.createElement(name)
        for key, value in data.items():
            if isinstance(value, dict):
                node = self.encode(value, key, document)
            else:
                node = document.createElement(key)
                node.appendChild(document.createTextNode(str(value)))
            parent_node.appendChild(node)
        return parent_node

    def save(self):
        """Update the configuration file based on the current settings."""
        xml_document = xml.dom.minidom.getDOMImplementation().createDocument(None, None, None)

        xml_document.appendChild(self.encode(self, 'configuration', xml_document))

        config_file = os.path.expanduser(self['data_dir'] + self['config_file'])
        if not os.path.exists(os.path.dirname(config_file)):
            os.makedirs(os.path.dirname(config_file))
        file = open(config_file, 'w')
        file.write(xml_document.toprettyxml())
        file.close()

class Managed_Window(object):
    """Interface to store and restore window position data."""
    # Defaults used when no other sources found
    x = 10
    y = 10
    width = 300
    height = 300
    _fullscreen = False
    
    def __init__(self, win_id, persistent_state=True):
        """Initialise with win_id used to identify instance of window."""
        self.id = win_id
        self.persistent_state = persistent_state
        # Set the title to the id.
        # This is likely to be overridded by subclasses but provides a
        # reasonable default
        self.set_title(win_id.capitalize())

        if self.id in config:
            if 'show' in config[self.id]:
                config[self.id]['show'] = True
            if 'x' in config[self.id]:
                self.x = config[self.id]['x']
            if 'y' in config[self.id]:
                self.y = config[self.id]['y']
            if 'width' in config[self.id]:
                self.width = config[self.id]['width']
            if 'height' in config[self.id]:
                self.height = config[self.id]['height']
            if 'fullscreen' in config[self.id]:
                self._fullscreen = config[self.id]['fullscreen']

        if self._fullscreen:
            self.maximize()
        else:
            screen_width = gtk.gdk.screen_width()
            screen_height = gtk.gdk.screen_height()
            width = self.width if self.width < screen_width else screen_width
            height = self.height if self.height < screen_height else screen_height
            self.move(int(self.x), int(self.y))
            self.set_default_size(width, height)

        self.connect('window-state-event', self.window_state_cb)
        self.connect("delete_event", self.delete_cb)

    def window_state_cb(self, widget, event):
        """Respond to a window state callback event."""
        self._fullscreen = bool(event.new_window_state & gtk.gdk.WINDOW_STATE_MAXIMIZED)

    def save_state(self):
        """Save the window position and size if different from default."""
        if self.persistent_state:
            if self.id not in config:
                config[self.id] = {} 
            config[self.id]['x'], config[self.id]['y'] = self.get_position()
            config[self.id]['width'], config[self.id]['height'] = self.get_size()
            config[self.id]['fullscreen'] = self._fullscreen
        elif self.id in config:
            if 'x' in config[self.id]:
                del config[self.id]['x']
            if 'y' in config[self.id]:
                del config[self.id]['y']
            if 'width' in config[self.id]:
                del config[self.id]['width']
            if 'height' in config[self.id]:
                del config[self.id]['height']
            if 'fullscreen' in config[self.id]:
                del config[self.id]['fullscreen']
            if not config[self.id].keys():
                del config[self.id]

    def destroy(self):
        """Destorys the window, saving the state first."""
        self.save_state()
        gtk.Widget.destroy(self)

    def delete_cb(self, widget, event):
        """Save the state before the window is deleted."""
        if 'show' in config[self.id]:
            config[self.id]['show'] = False
        self.save_state()
 
class Managed_Dialog(Managed_Window):
    """Interface to store and restore dialog position data."""
    
    def __init__(self, win_id, persistent_state=True):
        """Initialise with win_id used to identify instance of window."""
        Managed_Window.__init__(self, win_id, persistent_state)
        self.set_deletable(False)
        self.connect('response', self.save_state_cb)

    def save_state_cb(self, dialog, response):
        """Save the window position and size if different from default."""
        self.save_state()
        return False

class Allocation_Dialog(gtk.Dialog):
    """Simple dialog to change the allocation of an object."""

    def __init__(self, title, activity, x, y, callback):
        """Display the dialog."""
        gtk.Dialog.__init__(self, title,
                            buttons = (gtk.STOCK_CANCEL, gtk.RESPONSE_CANCEL,
                                       gtk.STOCK_OK, gtk.RESPONSE_OK))
        self.activity = activity
        self.initial_allocation = activity.allocation
        self.callback = callback
        self.connect('response', self.response_cb)
        self.move(int(x), int(y))
        slider = gtk.HScale()
        slider.set_range(0, 100)
        slider.set_value(self.initial_allocation * 100)
        self.last_value = slider.get_value()
        slider.connect('value-changed', self.allocation_changed_cb)
        self.vbox.pack_start(slider)
        self.show_all()

    def allocation_changed_cb(self, widget, scrolltype=None):
        """Update the allocation based on the widget value."""
        allocation = widget.get_value()
        self.activity.allocation = allocation / 100.0
        if self.last_value != allocation and (self.last_value == 0 or
                                              allocation ==0):
            if self.callback:
                self.callback()
        self.last_value = allocation

    def response_cb(self, dialog, response_id):
        """Process the response and close the dialog."""
        self.destroy()
        if response_id == gtk.RESPONSE_CANCEL:
            self.activity.allocation = self.initial_allocation
            if self.callback:
                self.callback()

class Timeline(gtk.Layout):
    """A widget used to represent events on a timeline."""
    x_pad = 5
    y_pad = 2
    draw_events = True
    draw_activities = True
    draw_divisions = True
    division_size = 0
    event_capture_size = 5
    editable = True
    zoom = 1.0

    class Marker(object):
        """Manages points on a timeline widget."""
        __slots__ = ['display', 'the_event', 'label', 'prior', 'subsequent', 'x_position']
        display = False

        def __init__(self, event, label = ""):
            self.the_event = event
            self.label = label
            self.x_position = -1
            self.subsequent = set()
            self.prior = set()

        def get_time(self):
            """Gets the time from external event."""
            if self.the_event:
                result = self.the_event.time
            else:
                result = datetime.datetime.now()
            return result

        def set_time(self, the_time):
            """Sets the time on the external event."""
            if self.the_event:
                self.the_event.time = the_time
            for activity in [item for item in self.prior]:
                if activity.start.time > the_time:
                    activity.swap()
            for activity in [item for item in self.subsequent]:
                if activity.end.time < the_time:
                    activity.swap()

        time = property(get_time, set_time)

    class Activity(object):
        """Manages activities on a timeline widget."""
        __slots__ = ['display', 'the_activity', 'start', 'end', 'label', 'y', 'count', 'session']
        display = True

        def __init__(self, activity, start_event, end_event, label, session):
            self.the_activity = activity
            self.start = start_event
            self.start.subsequent.add(self)
            self.end = end_event
            self.end.prior.add(self)
            self.label = label
            self.session = session
            self.y = 0
            self.count = 1

        def set_task(self, task):
            """Changes the task associated with this activity."""
            if self.the_activity:
                self.the_activity.set_task(task)
            else:
                self.the_activity = self.session.new_activity(task,
                                                              self.start.the_event,
                                                              self.end.the_event)
            self.label = task.name

        def delete(self):
            """Removes the activity."""
            self.start.subsequent.remove(self)
            if self.end:
                self.end.prior.remove(self)
            if self.the_activity:
                self.the_activity.delete()
                self.the_activity = None

        def get_duration(self):
            """Returns a string containing the duration of this activity."""
            duration = self.end.time - self.start.time
            return timedelta_to_string(duration)

        def swap(self):
            """Swaps the start and end events for this activity."""
            self.start, self.end = self.end, self.start
            self.start.subsequent.add(self)
            self.start.prior.remove(self)
            self.end.prior.add(self)
            self.end.subsequent.remove(self)
            if self.the_activity:
                self.the_activity.swap()

        def change_allocation(self, x, y, callback):
            """Popup the window used to change the activity time allocation."""
            if self.the_activity and self.the_activity.task:
                title = 'Change Allocation for %s' % (self.the_activity.task.full_name)
                Allocation_Dialog(title, self.the_activity, x, y, callback)

    def __init__(self, date, session=None, callback=None, parent=None):
        """Initialise the timeline widget."""
        if parent:
            hadjustment = parent.get_hadjustment()
            vadjustment = parent.get_vadjustment()
        else:
            hadjustment = None
            vadjustment = None
        gtk.Layout.__init__(self, hadjustment, vadjustment)
        self.connect("expose_event", self.expose)
        self.connect("motion_notify_event", self.mouse_motion_cb)
        self.connect("button_press_event", self.mouse_button_cb)
        self.connect("button_release_event", self.mouse_button_release_cb)
        self.add_events(gtk.gdk.POINTER_MOTION_MASK | gtk.gdk.BUTTON_PRESS_MASK | gtk.gdk.BUTTON_RELEASE_MASK)
        self._parent = parent
        self.callback = callback
        self.ratio = 0
        self.height = 10
        self.markers = {}
        self.activities = []
        self.captured_event = None
        self.captured_activity = None
        self.drag_event = False
        self.drag_activity = False
        self.click_x = None
        self.session = session
        self.children = []
        self.gradient = {}
        self.__old_height = self.height - 1
        self.date = date
        if parent:
            parent.children.append(weakref.ref(self))
        start = datetime.time(hour = int(config.get('day_start')[:2]),
                              minute = int(config.get('day_start')[3:]),
                              second = 0,
                              microsecond = 0)
        end = datetime.time(hour = int(config.get('day_end')[:2]),
                            minute = int(config.get('day_end')[3:]),
                            second = 0,
                            microsecond = 0)
        self.start = datetime.datetime.combine(date, start)
        self.end = datetime.datetime.combine(date, end)

    def get_ratio(self):
        """Returns the ratio for this timeline."""
        if self._parent:
            return self._parent.ratio
        else:
            return self._ratio

    def set_ratio(self, ratio):
        """Sets the ratio for the group."""
        if self._parent:
            self._parent.ratio = ratio
        else:
            self._ratio = ratio

    ratio = property(get_ratio, set_ratio)

    def create_gradients(self):
        """Create the colour gradients used to draw the activities."""
        if self.__old_height != self.height:
            self.__old_height = self.height
        else:
            return
        gradients = {'yellow' : ((0.5, 0.5, 0.0), (1.0, 1.0, 0.5)),
                     'red'    : ((0.5, 0.0, 0.0), (1.0, 0.5, 0.5)),
                     'green'  : ((0.0, 0.5, 0.0), (0.5, 1.0, 0.5)),
                     'blue'   : ((0.0, 0.0, 0.5), (0.5, 0.5, 1.0))}
        
        for colour in gradients.keys():
            for alpha in [0.9, 0.5]:
                i = '%s_%02.02f' % (colour, alpha)
                self.gradient[i] = cairo.LinearGradient(0, 0, 0, self.height)
                red, green, blue = gradients[colour][0]
                self.gradient[i].add_color_stop_rgba(0, red, green, blue, alpha)
                red, green, blue = gradients[colour][1]
                self.gradient[i].add_color_stop_rgba(1, red, green, blue, alpha)
                
                i = '%s_%02.02f_selected' % (colour, alpha)
                self.gradient[i] = cairo.LinearGradient(0, 0, 0, self.height)
                red, green, blue = gradients[colour][0]
                self.gradient[i].add_color_stop_rgba(1, red, green, blue, alpha)
                red, green, blue = gradients[colour][1]
                self.gradient[i].add_color_stop_rgba(0, red, green, blue, alpha)

        self.light_gradient = cairo.LinearGradient(0, 0, 0, self.height)
        self.light_gradient.add_color_stop_rgba(0.0, 1.0, 1.0, 1.0, 0.25)
        self.light_gradient.add_color_stop_rgba(1.0, 1.0, 1.0, 1.0, 0.5)

    def add_child(self, child):
        """Adds a weak reference to the child widget."""
        self.children.append(weakref.ref(child))

    def make_event_menu(self, event):
        """Creates the context menu for events"""
        event_menu = [('Merge Events', self.merge_events_cb)]
        if not event.prior:
            event_menu.append(('< Expand Left', self.expand_cb))
        elif not event.subsequent:
            event_menu.append(('> Expand Right', self.expand_cb))
        else:
            event_menu.append(('Merge Activities', self.merge_cb))
        if len(event.prior) + len(event.subsequent) > 1:
            event_menu.append(('Split Event', self.split_event_cb))
        return build_menu(event_menu, False)

    def make_task_menu(self, activity):
        """Creates the context menu for events"""
        task_menu = []
        if self.session.task_manager:
            task_menu = self.session.task_manager.create_menu_definition(self.start_task_cb,
                                                                 config.get('hierarchical_list'))

        if self.session.current_activity or self.session.previous_task:
            task_menu.append(('---', '---'))

        if self.session.current_activity:
            task_menu.append(('Stop %s' % (self.session.current_activity.task.name), self.stop_current_activity_cb))

        if self.session.previous_task:
            if (not self.session.current_activity or
                self.session.previous_task != self.session.current_activity.task):
                task_menu.append(('Resume %s' % (self.session.previous_task.name),
                                  self.resume_task_cb))

        return build_menu(task_menu, False)

    def make_activity_menu(self, activity):
        """Creates the context menu for events"""
        activity_menu = []
        if activity.the_activity and activity.the_activity.is_active():
            activity_menu.extend([('Stop Activity', self.stop_activity_cb),
                                  ('---', '---')])
        if self.session.task_manager:
            task_menu = self.session.task_manager.create_menu_definition(self.change_task_cb,
                                                            config.get('hierarchical_list'))
            activity_menu.extend([('>>>Set Task', task_menu)])
        activity_menu.extend([('Occupy Full Day', self.all_day_activity_cb),
                              ('Split Activity', self.split_activity_cb),
                              ('Allocate 100%', self.make_singular_activity_cb)])
        if activity.the_activity:
            activity_menu.append(('Change Allocation', self.change_allocation_cb))
        if activity.end.the_event:
            activity_menu.extend([('Delete Activity', self.delete_activity_cb)])
        shift_menu = [('---', '---')]
        if not activity.start.prior or not activity.end.subsequent:
            activity_menu.extend([('---', '---'),
                                  ('<> Grow', self.grow_cb, 'both')])
        if not activity.start.prior:
            activity_menu.extend([('<] Grow Left', self.grow_cb, 'left')])
            shift_menu.extend([('<< Shift Left', self.shift_cb, 'left')])
        if not activity.end.subsequent:
            activity_menu.extend([('[> Grow Right', self.grow_cb, 'right')])
            shift_menu.extend([('>> Shift Right', self.shift_cb, 'right')])
        if len(shift_menu) > 1:
            activity_menu.extend(shift_menu)

        return build_menu(activity_menu, False)

    def expose(self, widget, event):
        """Updates display of widget."""
        self.height = self.allocation.height

        self.create_gradients()
        self.context = widget.bin_window.cairo_create()
        self.context.select_font_face('sans-serif',
                                      cairo.FONT_SLANT_NORMAL,
                                      cairo.FONT_WEIGHT_NORMAL)
        self.context.rectangle(event.area.x, event.area.y,
                               event.area.width, event.area.height)
        self.context.clip()
        self.draw(self.context)
        return False

    def mouse_button_cb(self, widget, event):
        """Processes the press of a mouse button on a timeline."""
        if event.button == 1:
            if self.editable:
                if self.captured_event:
                    self.drag_event = True
                elif self.captured_activity:
                    if (not self.captured_activity.the_activity or
                        not self.captured_activity.the_activity.is_active()):
                        self.drag_activity = True
                        self.drag_start = event.x
                        self.queue_draw()
                else:
                    self.click_x = event.x
        elif event.button == 2:
            pass
        elif event.button == 3:
            self.x = event.x
            if self.captured_event:
                menu = self.make_event_menu(self.captured_event)
            elif self.captured_activity:
                menu = self.make_activity_menu(self.captured_activity)
            elif self.editable:
                menu = self.make_task_menu(self.captured_activity)
            else:
                menu = None
            if menu:
                menu.popup(None, None, None, event.button, event.time)

    def mouse_button_release_cb(self, widget, event):
        """Processes the release of a mouse button on a timeline."""
        self.click_x = None
        if self.drag_event:
            capture_size = self.event_capture_size
            for time_event in self.markers.values():
                if abs(event.x - time_event.x_position) < capture_size:
                    cursor = gtk.gdk.Cursor(gtk.gdk.SB_H_DOUBLE_ARROW)
                    widget.window.set_cursor(cursor)
                    self.captured_event = time_event
                    capture_size = abs(event.x - time_event.x_position)
            if not self.captured_event:
                widget.window.set_cursor(None)
        self.drag_event = False
        self.drag_activity = False
        self.mouse_motion_cb(widget, event)
        if self.callback:
            self.callback()
        self.queue_draw()

    def move_activity(self, activity, time_delta):
        """Move the activitiy by the specified time."""
        max_delta = activity.end.time.replace(hour = 23,
                                              minute = 59,
                                              second = 59) - activity.end.time
        if activity.the_activity:
            end_event = activity.the_activity.end_event
            if self.session.current_activity in end_event.subsequent_activity:
                max_delta = datetime.datetime.now() - activity.end.time
        if time_delta > max_delta:
            time_delta = max_delta
        min_delta = activity.start.time.replace(hour = 0,
                                                minute = 0,
                                                second = 0) - activity.start.time
        if time_delta < min_delta:
            time_delta = min_delta
        activity.start.time += time_delta
        activity.end.time += time_delta
        self.normalise_event(activity.start)
        self.normalise_event(activity.end)

    def move_event(self, event, time_delta):
        """Moves the event by the specified time."""
        time = datetime.datetime.combine(self.date, self._parent.start.time())
        time += time_delta
        if self.session.current_activity in event.the_event.subsequent_activity:
            max_time = datetime.datetime.now()
        else:
            max_time = event.time.replace(hour = 23, minute = 59, second = 59, microsecond = 99999)
        if time > max_time:
            time = max_time
        min_time = event.time.replace(hour = 0, minute = 0, second = 0, microsecond = 0)
        if time < min_time:
            time = min_time
        event.time = time

    def mouse_motion_cb(self, widget, event):
        """Processes the mouse movement on a timeline."""
        self.x = event.x_root
        self.y = event.y_root
        if self.click_x and self.click_x != self.x:
            time = self.click_x / self.ratio
            start_event = self.session.new_event()
            day_start = datetime.datetime.combine(self.date, self._parent.start.time())
            start_event.time = day_start + datetime.timedelta(hours = time, seconds = 0)
            end_event = self.session.new_event()
            end_event.time = day_start + datetime.timedelta(hours = time, seconds = 1)
            activity = self.new_activity(None, start_event, end_event)
            self.captured_event = activity.end
            self.drag_event = True
            self.queue_draw()
            self.click_x = None
        if self.drag_event:
            time = event.x / self.ratio
            self.move_event(self.captured_event, datetime.timedelta(hours = time))
            self.normalise_event(self.captured_event)
            if self.captured_event.label:
                test = "%s: %s" % (label, self.captured_event.time.strftime('%H:%M:%S'))
            else:
                text = "Time: %s" % (self.captured_event.time.strftime('%H:%M:%S'))
            self.set_tooltip_text(text)
            self.update_start_end()
            self.queue_draw()
        elif self.drag_activity:
            x_delta = event.x - self.drag_start
            self.drag_start = event.x
            delta_hours = x_delta / self.ratio
            timedelta = datetime.timedelta(hours = delta_hours)
            self.move_activity(self.captured_activity, timedelta)
            self.update_start_end()
            self.queue_draw()
        elif self.editable:
            self.captured_event = None
            self.captured_activity = None
            capture_size = self.event_capture_size
            for time_event in self.markers.values():
                if (abs(event.x - time_event.x_position) < capture_size and
                    time_event.the_event and time_event.the_event.get_id() != 0):
                    cursor = gtk.gdk.Cursor(gtk.gdk.SB_H_DOUBLE_ARROW)
                    widget.window.set_cursor(cursor)
                    self.captured_event = time_event
                    capture_size = abs(event.x - time_event.x_position)
                    if self.captured_event.label:
                        label = "%s: %s" % (self.captured_event.label, self.captured_event.time.strftime('%H:%M:%S'))
                    else:
                        label = "Time: %s" % (self.captured_event.time.strftime('%H:%M:%S'))
                    if config.get('verbosity') & 8:
                        label = "%d: %s" % (self.captured_event.the_event.get_id(),
                                            label)
                    self.set_tooltip_text(label)
            if not self.captured_event:
                widget.window.set_cursor(None)
                for activity in self.activities:
                    if event.x > activity.start.x_position and event.x < activity.end.x_position:
                        height = (self.height - 2 * self.y_pad) / activity.count
                        y = self.y_pad + activity.y * height
                        if event.y > y and event.y < y + height:
                            cursor = gtk.gdk.Cursor(gtk.gdk.HAND1)
                            widget.window.set_cursor(cursor)
                            self.captured_activity = activity
                            if self.captured_activity.the_activity:
                                if config.get('verbosity') & 8:
                                    label = "%d: %s" % (self.captured_activity.the_activity.get_id(),
                                                        self.captured_activity.the_activity.task.full_name())
                                else:
                                    label = self.captured_activity.the_activity.task.full_name()
                            else:
                                label = self.captured_activity.label
                            duration = self.captured_activity.get_duration()
                            self.set_tooltip_text("%s\n%s" % (label, duration))
            else:
                if self.captured_event.label:
                    label = "%s: %s" % (self.captured_event.label, self.captured_event.time.strftime('%H:%M:%S'))
                else:
                    label = "Time: %s" % (self.captured_event.time.strftime('%H:%M:%S'))
                if config.get('verbosity') & 8:
                    label = "%d: %s" % (self.captured_event.the_event.get_id(),
                                        label)
                self.set_tooltip_text(label)
        return False

    def change_task_cb(self, event, task):
        """Changes the task associated with an activity."""
        if self.captured_activity:
            self.captured_activity.set_task(task)
            self.queue_draw()
            if self.callback:
                self.callback()

    def start_task_cb(self, event, task):
        """Starts a new activity, stopping the current if necessary."""
        self.session.start_task(task)

        self.queue_draw()
        if self.callback:
            self.callback()

    def stop_activity_cb(self, event):
        """Stops the selected activity."""
        if self.captured_activity:
            self.stop_activity(self.captured_activity)

            self.queue_draw()
            if self.callback:
                self.callback()

    def stop_activity(self, activity):
        """Stops the activity."""
        if activity.the_activity and activity.the_activity.is_active():
            self.session.pause() # This will trigger a re-draw of the timesheet.
        else:
            Message.show('Invalid action',
                         'Attempt to stop non active task: %s' % (activity.label),
                         gtk.MESSAGE_WARNING)

    def stop_current_activity_cb(self, event):
        """Stops the current activity."""
        if self.session.current_activity:
            self.session.pause()

            self.queue_draw()
            if self.callback:
                self.callback()

    def resume_task_cb(self, event):
        """Resumes the previous activity."""
        if self.session.previous_task:
            self.session.resume()

            self.queue_draw()
            if self.callback:
                self.callback()

    def split_activity_cb(self, item):
        """Splits an activity in 2"""
        if self.captured_activity and self.editable:

            day_start = datetime.datetime.combine(self.date,
                                                  self._parent.start.time())
            time = self.x / self.ratio
            new_event = self.session.new_event()
            new_event.time = day_start + datetime.timedelta(hours = time)

            self.split_activity(self.captured_activity, new_event)

            self.queue_draw()
            del self.x
            if self.callback:
                self.callback()

    def split_activity(self, activity, event):
        """Splits an activity in 2"""
        start_event = activity.start.the_event
        new_activity = None
        if activity.the_activity:
            new_activity = activity.the_activity.split(event)
        new_gui_activity = self.new_activity(new_activity, start_event, event)
        self.change_start(activity, new_gui_activity.end)
        return new_gui_activity

    def all_day_activity_cb(self, item):
        """Change the duration of the activity to full day duration."""
        if self.captured_activity and self.editable:
            self.occupy_full_day(self.captured_activity)

            self.queue_draw()
            del self.x
            if self.callback:
                self.callback()

    def occupy_full_day(self, activity):
        """Makes the activity last all day."""
        if activity.the_activity:
            activity.the_activity.occupy_full_day()
        else:
            start = datetime.time(hour = int(config.get('day_start')[:2]),
                                  minute = int(config.get('day_start')[3:]),
                                  second = 0,
                                  microsecond = 0)
            start = datetime.datetime.combine(self.date, start)
            duration = string_to_timedelta(config['hours_per_week'] / config['days_per_week'])
            end = start + duration
            if not activity.start.prior:
                activity.start.time = start
            else:
                new_event = self.session.new_event()
                new_event.time = start
                self.change_start(activity, self.new_marker(new_event))
                
            if not activity.end.subsequent:
                activity.end.time = end
            else:
                new_event = self.session.new_event()
                new_event.time = end
                self.change_end(activity, self.new_marker(new_event))
                

    def make_singular_activity_cb(self, item):
        """Makes the current activity the single thread."""
        if self.captured_activity and self.editable:
            self.make_singular_activity(self.captured_activity)

            self.queue_draw()
            del self.x
            if self.callback:
                self.callback()

    def make_singular_activity(self, activity):
        """Makes the activity the single thread."""
        delete_activities = set()

        for other_activity in self.activities:
            if other_activity == activity:
                continue

            if (other_activity.start.time >= activity.start.time and
                other_activity.end.time <= activity.end.time):
                delete_activities.add(other_activity)

            elif (other_activity.start.time <= activity.start.time and
                  other_activity.end.time >= activity.end.time and activity.end.the_event):

                new_activity = self.split_activity(other_activity, activity.start.the_event)
                if new_activity.start.time == activity.start.time:
                    delete_activities.add(new_activity)

                if other_activity.end.time == activity.end.time:
                    delete_activities.add(other_activity)
                else:
                    self.change_start(other_activity, activity.end)

            elif (other_activity.start.time <= activity.start.time and
                  other_activity.end.time >= activity.start.time):
                self.change_end(other_activity, activity.start)

            elif (other_activity.start.time <= activity.end.time and
                  other_activity.end.time >= activity.end.time and activity.end.the_event):
                self.change_start(other_activity, activity.end)

        for activity in delete_activities:
            self.delete_activity(activity)

    def change_allocation_cb(self, item):
        """Changes the allocation of this activity."""
        if self.captured_activity and self.editable:
            activity = self.captured_activity
            activity.change_allocation(self.x, self.y, self.callback)
            self.queue_draw()

    def delete_activity_cb(self, event):
        """Deletes the currently selected activity."""
        if self.captured_activity:
            self.delete_activity(self.captured_activity)

            self.captured_activity = None

            self.queue_draw()
            if self.callback:
                self.callback()

    def delete_activity(self, activity):
        """Deletes the currently selected activity."""
        activity.label = None
        if activity.the_activity and activity.the_activity.is_active():
            self.session.current_activity = None
        self.activities.remove(activity)
        activity.delete()
        self.delete_event(activity.start)
        self.delete_event(activity.end)

    def grow_cb(self, event, direction):
        """Expands the current activity in a given direction."""
        if self.captured_activity:
            self.grow(self.captured_activity, direction)

            self.queue_draw()
            if self.callback:
                self.callback()

    def grow(self, activity, direction):
        """Expands an activity in a given direction."""
        if direction == 'left' or direction == 'both':
            self.expand(activity.start)
        if direction == 'right' or direction == 'both':
            self.expand(activity.end)

    def shift_cb(self, event, direction):
        """Moves the current activity in a given direction."""
        if self.captured_activity:
            self.shift(self.captured_activity, direction)

            self.queue_draw()
            if self.callback:
                self.callback()

    def shift(self, activity, direction):
        """Moves an activity in a given direction."""
        if direction == 'left':
            time = activity.start.time
            self.expand(activity.start)
            delta = time - activity.start.time
            activity.end.time -= delta

        if direction == 'right':
            time = activity.end.time
            self.expand(activity.end)
            delta = activity.end.time - time
            activity.start.time += delta

            # if the end is fixed (eg the current task) don't let the start move past it
            duration = activity.end.time - activity.start.time
            constrained = False
            for activity in activity.start.subsequent:
                if activity.start.time > activity.end.time:
                    activity.start.time = activity.end.time
                    constrained = True
            if constrained:
                activity.end.time = activity.start.time + duration

    def expand_cb(self, event):
        """Moves the current event to the previous / next event based on whether this is a start or stop event."""
        if self.captured_event:
            if self.expand(self.captured_event):
                self.captured_event = None
            self.queue_draw()
            if self.callback:
                self.callback()

    def expand(self, event):
        """Moves an event to the previous / next event based on whether this is a start or stop event."""
        events = self.markers.values()
        events.sort(key=operator.attrgetter('time'))
        index = events.index(event)
        event_changed = False
        if event.prior and not event.subsequent and index < len(events) - 1:
            for activity in [activity for activity in event.prior]:
                self.change_end(activity, events[index + 1])
            event_changed = True
        elif event.subsequent and not event.prior and index > 0:
            prior_event = events[index - 1] if events[index - 1].the_event else events[index - 2]
            for activity in [activity for activity in event.subsequent]:
                self.change_start(activity, prior_event)
            event_changed = True
        elif index == 0:
            work_day_start = event.time.replace(hour = int(config.get('day_start')[:2]),
                                                minute = int(config.get('day_start')[3:]),
                                                second = 0, microsecond = 0)
            if event.time > work_day_start:
                event.time = work_day_start
            else:
                event.time = event.time.replace(hour = 0,
                                                minute = 0,
                                                second = 0,
                                                microsecond = 0)
            event_changed = True
        elif index == len(events) - 1:
            work_day_end = event.time.replace(hour = int(config.get('day_end')[:2]),
                                              minute = int(config.get('day_end')[3:]),
                                              second = 0,
                                              microsecond = 0)
            if event.time < work_day_end:
                event.time = work_day_end
            else:
                event.time = event.time.replace(hour = 23,
                                                minute = 59,
                                                second = 59,
                                                microsecond = 0)
            event_changed = True

        return event_changed

    def merge_events_cb(self, event):
        """Merge all events within the time window of the captured event."""
        if self.captured_event:
            self.merge_events(self.captured_event)

            self.queue_draw()
            if self.callback:
                self.callback()

    def merge_events(self, event):
        """Merge all events within the time window of the event."""
        for other_event in self.markers.values():
            if other_event != event:
                delta = abs(get_decimal_hours(event.time - other_event.time)) * 60.0
                if delta < config.get('merge_event_window'):
                    for activity in [prior for prior in other_event.prior]:
                        self.change_end(activity, event)

                    for activity in [subsequent for subsequent in other_event.subsequent]:
                        self.change_start(activity, event)

    def merge_cb(self, event):
        """Deletes the currently selected activity."""
        if self.captured_event:
            self.merge(self.captured_event)

            self.queue_draw()
            if self.callback:
                self.callback()

    def merge(self, event):
        """Deletes the currently selected activity."""
        end_event = None
        for subsequent in event.subsequent:
            if subsequent.end.the_event:
                if not end_event or subsequent.end < end_event:
                    end_event = subsequent.end
                    del_activity = subsequent

        if end_event:
            for activity in [prior for prior in event.prior]:
                self.change_end(activity, end_event)

            self.delete_activity(del_activity)

    def split_event_cb(self, event):
        """Assigns a new event to each of the activities starting and stopping at this event."""
        if self.captured_event and self.editable:
            self.split_event(self.captured_event)

    def split_event(self, event):
        """Assigns a new event to each of the activities starting and stopping at this event."""
        for activity in [activity for activity in event.prior]:
            new_event = self.session.new_event()
            new_event.time = event.time
            new_marker = self.new_marker(new_event)
            self.change_end(activity, new_marker)

        for activity in [activity for activity in event.subsequent]:
            new_event = self.session.new_event()
            new_event.time = event.time
            new_marker = self.new_marker(new_event)
            self.change_start(activity, new_marker)

    def normalise_event(self, event):
        """If activity start time after end time swap event times."""
        swap = []
        for prior in event.prior:
            if prior.start.time > event.time:
                swap.append(prior)
        for subsequent in event.subsequent:
            if subsequent.end.the_event:
                if subsequent.start.time > subsequent.end.time:
                    swap.append(subsequent)
            elif event.time > subsequent.end.time:
                event.time = subsequent.end.time

        for activity in swap:
            activity.swap()

    def rounded_rectangle(self, context, x, y, width, height, radius):
        """Creates a path for a rectangle with rounded corners."""
        x1 = x + width
        y1 = y + height
        if width / 2 < radius:
            if height / 2 < radius:
                context.move_to(x, (y + y1) / 2)
                context.curve_to(x ,y, x, y, (x + x1) / 2, y)
                context.curve_to(x1, y, x1, y, x1, (y + y1) / 2)
                context.curve_to(x1, y1, x1, y1, (x1 + x) / 2, y1)
                context.curve_to(x, y1, x, y1, x, (y + y1) / 2)
            else:
                context.move_to(x, y + radius)
                context.curve_to(x ,y, x, y, (x + x1) / 2, y)
                context.curve_to(x1, y, x1, y, x1, y + radius)
                context.line_to(x1 , y1 - radius)
                context.curve_to(x1, y1, x1, y1, (x1 + x) / 2, y1)
                context.curve_to(x, y1, x, y1, x, y1- radius)
        else:
            if height / 2 < radius:
                context.move_to(x, (y + y1) / 2)
                context.curve_to(x , y, x , y, x + radius, y)
                context.line_to(x1 - radius, y)
                context.curve_to(x1, y, x1, y, x1, (y + y1) / 2)
                context.curve_to(x1, y1, x1, y1, x1 - radius, y1)
                context.line_to(x + radius, y1)
                context.curve_to(x, y1, x, y1, x, (y + y1) / 2)
            else:
                context.move_to(x, y + radius)
                context.curve_to(x , y, x, y, x + radius, y)
                context.line_to(x1 - radius, y)
                context.curve_to(x1, y, x1, y, x1, y + radius)
                context.line_to(x1 , y1 - radius)
                context.curve_to(x1, y1, x1, y1, x1 - radius, y1)
                context.line_to(x + radius, y1)
                context.curve_to(x, y1, x, y1, x, y1- radius)

        context.close_path()

    def draw_activity(self, context, x, y, width, height, text, gradient1, gradient2, text_colour, min_x, max_x):
        """Displays a representation of an activity on the timeline."""
        radius = 10.0;

        if not width or not height:
            return

        context.translate(0.0, y)
        context.scale(1.0, float(height) / float(self.height - 2 * self.y_pad))
        self.rounded_rectangle(context, x, 0, width, self.height - 2 * self.y_pad, radius)
        context.set_source(gradient1)
        context.fill_preserve()
        context.set_source_rgba(0.0, 0.0, 0.5, 0.5)
        context.set_line_width(1.0)
        context.stroke()

        context.identity_matrix()
        if text:
            context.set_source_rgb(text_colour[0], text_colour[1], text_colour[2])
            context.save()
            context.new_path()
            context.rectangle(x, y, width, height)
            context.clip()
            extents = context.text_extents(text)
            if width > extents[2]:
                text_x = x + 0.5 * (width - extents[2])
            else:
                text_x = x
            if text_x < min_x and min_x < x + width:
                text_x = min_x
            if text_x + extents[2] > max_x:
                text_x = max(max_x - extents[2], x)
            text_y = y + 0.5 * (height + extents[3])

            context.move_to(text_x, text_y)
            context.show_text(text)
            context.stroke()
            context.fill()
            context.restore()

        context.translate(0.0, y)
        context.scale(1.0, float(height) / float(self.height - 2 * self.y_pad))
        self.rounded_rectangle(context, x + self.x_pad, self.y_pad, width - 2 * self.x_pad, (self.height - 2 * self.y_pad) / 2 - self.y_pad, radius - self.x_pad / 2)
        context.set_source(gradient2)
        context.fill()
        context.identity_matrix()

    def draw_event(self, context, x, y, width, height, text, gradient1, gradient2):
        """Displays a representation of an activity on the timeline."""
        radius = 10.0;

        if not width or not height:
            return

        self.rounded_rectangle(context, x, y, width, height, radius)
        context.set_source(gradient1)
        context.fill_preserve()
        context.set_source_rgba(0.0, 0.0, 0.5, 0.5)
        context.set_line_width(1.0)
        context.stroke()

        if text:
            context.set_source_rgb(0, 0, 0)
            extents = context.text_extents(text)
            text_x = x + 0.5 * (width + extents[3])
            text_y = 0.5 * (self.height + extents[2])

            context.move_to(text_x, text_y)
            context.rotate(3.0 * math.pi / 2.0)
            context.show_text(text)
            context.stroke()
            context.identity_matrix()

        self.rounded_rectangle(context,
                               x + self.x_pad,
                               y + self.y_pad,
                               width - 2 * self.x_pad,
                               height / 2 - self.y_pad,
                               radius - self.x_pad / 2)
        context.set_source(gradient2)
        context.fill()
        context.identity_matrix()

    def get_division_size(self):
        """Determines an appropriate size of the display of divisions."""
        times = [60.0, 30.0, 20.0, 15.0, 10.0, 5.0, 1.0]
        times.sort()
        times.reverse()
        ratio = self.ratio
        for division_size in times:
            if ratio * division_size / 60.0 < config.get('division_spacing'):
                break
        return division_size

    def check_width(self):
        """Updates the width of the widget as necessary to fit all events."""
        if self._parent:
            width = self._parent.check_width()
        else:
            old_width = int(get_decimal_hours(self.end - self.start) * self.ratio + 0.5)

            start = datetime.time(hour = int(config.get('day_start')[:2]),
                                       minute = int(config.get('day_start')[3:]),
                                       second = 0,
                                       microsecond = 0)
            end = datetime.time(hour = int(config.get('day_end')[:2]),
                                     minute = int(config.get('day_end')[3:]),
                                     second = 0,
                                     microsecond = 0)

            for child in self.children:
                if child():
                    if child().start.time() < start:
                        start = child().start.time()
                    if child().end.time() > end:
                        end = child().end.time()

            self.start = datetime.datetime.combine(self.date, start)
            self.end = datetime.datetime.combine(self.date, end)
            width = int(get_decimal_hours(self.end - self.start) * self.ratio + 0.5)

            adjustment = self.get_hadjustment()
            if width != old_width:
                width = int(max(width, adjustment.upper) + 0.5)
                self.set_size(width, int(self.height + 0.5))
                self.queue_draw()
                for child in self.children:
                    if child():
                        child().set_size(width, int(self.height + 0.5))
                        child().queue_draw()
            width = int(max(width, adjustment.upper) + 0.5)
        return width

    def update_start_end(self):
        """Updates the start and end time for this timeline."""
        events = self.markers.values()
        if events:
            day_start = datetime.datetime.combine(self.date, datetime.time())
            day_end = datetime.datetime.combine(self.date, datetime.time(23, 59, 59))
            margin = datetime.timedelta(minutes = self.get_division_size())
            events.sort(key=operator.attrgetter('time'))
            self.start = events[0].time
            if self.start - margin > day_start:
                self.start -= margin
            else:
                self.start = day_start

            self.end = events[-1].time + margin
            if self.end + margin < day_end:
                self.end += margin
            else:
                self.end = day_end

    def draw(self, context):
        """Draws the timeline widget."""
        if not self.ratio:
            self.calculate_ratio(context)

        width = self.check_width()
        self.calculate_heights()

        adjustment = self.get_hadjustment()
        min_x = adjustment.value
        max_x = min_x + adjustment.page_size

        if self._parent:
            start = datetime.datetime.combine(self.date,
                                              self._parent.start.time())
            end = datetime.datetime.combine(self.date,
                                            self._parent.end.time())
        else:
            start = self.start
            end = self.end

        context.save()
        context.reset_clip()
        self.bin_window.clear()
        context.restore()

        if self.draw_divisions:
            division_size = datetime.timedelta(minutes = self.get_division_size())
            # Determine start time rounded up to division size
            division = datetime.datetime.combine(self.date, datetime.time())
            context.set_source_rgba(0.0, 0.0, 0.0, 0.3);
            while division < datetime.datetime.combine(self.date,
                                                       datetime.time(23, 59, 59)):
                time = get_decimal_hours(division - start)
                x = time * self.ratio
                if x > min_x and x < max_x:
                    text = division.strftime('%H:%M')
                    extents = context.text_extents(text)
                    context.move_to(x, 0)
                    context.line_to(x, (self.height - extents[2]) * 0.5)
                    context.move_to(x, (self.height + extents[2]) * 0.5)
                    context.line_to(x, self.height)
                    context.move_to(x + extents[3] * 0.5, (self.height + extents[2]) * 0.5)
                    context.rotate(3.0 * math.pi / 2.0)
                    context.show_text(text)
                    context.identity_matrix()
                division += division_size
            context.set_source_rgba(0.0, 0.0, 0.0, 0.3);
            context.stroke()

        # Draw an indication of the current time
        now = datetime.datetime.now()
        if now > start and now < end:
            x = get_decimal_hours(now - start) * self.ratio
            context.move_to(x, 0)
            context.line_to(x, self.height)
            context.set_source_rgb(1.0, 0.3, 0.3);
            context.stroke()

        alpha = 0.75
        context.move_to(0, 0)
        context.line_to(width, 0)
        context.set_source_rgba(0.3, 0.3, 0.3, alpha);
        context.stroke()
        context.move_to(0, self.height)
        context.line_to(width, self.height)
        context.set_source_rgba(1.0, 1.0, 1.0, alpha);
        context.stroke()

        active = []
        start_conflict = 0

        events = self.markers.keys()
        events.sort(key=operator.attrgetter('time'))
        for event in events:
            # Check if the active task has been completed
            if 0 and event.get_id() == 0:
                for activity in self.markers[event].prior:
                    if activity.the_activity and activity.the_activity.end_event:
                        event.the_event = activity.the_activity.end_event
                for activity in self.markers[event].subsequent:
                    if activity.the_activity and activity.the_activity.start_event:
                        event.the_event = activity.the_activity.start_event
            time = get_decimal_hours(event.time - start)
            x = time  * self.ratio
            self.markers[event].x_position = x

            active.extend(self.markers[event].subsequent)
            for activity in self.markers[event].prior:
                active.remove(activity)
            count = 0
            for activity in active:
                if activity.the_activity:
                    task = activity.the_activity.task
                    if task:
                        if not task.booking_number or task.get_booking_number()[0] != '#':
                            count += 1
                else:
                    count += 1
            if not start_conflict and count > 1:
                start_conflict = x
            if count < 2 and start_conflict:
                context.rectangle(start_conflict, 0, x - start_conflict, self.height);
                context.set_source_rgba(1.0, 0, 0, 0.3);
                context.fill()
                start_conflict = 0

            if (self.draw_events and event and event.get_id() != 0 and
                x + self.x_pad > min_x and x - self.x_pad < max_x):
                gradient = self.gradient['green_0.90']
                gradient2 = self.light_gradient
                self.draw_event(context, x - self.x_pad / 2, 0, self.x_pad, self.height, event.label,
                                gradient, gradient2)

        if self.draw_activities:
            for activity in self.activities:
                time = get_decimal_hours(activity.start.time - start)
                duration = get_decimal_hours(activity.end.time - activity.start.time)
                width = duration * self.ratio
                x = time  * self.ratio
                height = (self.height - 2 * self.y_pad) / activity.count
                y = self.y_pad + activity.y * height

                gradient2 = self.light_gradient
                if activity.the_activity and activity.the_activity.task.get_booking_number():
                    if activity.the_activity.task.get_booking_number()[0] == '#':
                        text_colour = (0, 0, 0)
                        colour = 'green'
                    else:
                        text_colour = (1, 1, 1)
                        colour = 'blue'
                    if activity.the_activity.allocation == 0:
                        colour = 'red'
                else:
                    text_colour = (0, 0, 0)
                    colour = 'yellow'
                
                clip = False
                if config.get('include_future'):
                    alphas = [0.9]
                elif activity.start.time > now:
                    alphas = [0.5]
                elif activity.end.time < now:
                    alphas = [0.9]
                else:
                    clip = True
                    alphas = [0.5, 0.9]
                    
                for alpha in alphas:
                    if self.drag_activity and (self.captured_activity == activity):
                        i = '%s_%02.02f_selected' % (colour, alpha)
                    else:
                        i = '%s_%02.02f' % (colour, alpha)
                    gradient = self.gradient[i]
                    gradient2 = self.light_gradient
    
                    if activity.the_activity:
                        text = activity.the_activity.task.name
                    else:
                        text = activity.label
                    context.save()
                    if clip:
                        time = get_decimal_hours(now - start)
                        x_now = time * self.ratio
                        if alpha == 0.5:
                            context.rectangle(x_now, 0, width, self.height)
                            context.clip()
                        else:
                            context.rectangle(0, 0, x_now, self.height)
                            context.clip()
                    self.draw_activity(context, x, y, width, height, text,
                                       gradient, gradient2, text_colour,
                                       min_x, max_x)
                    context.restore()

    def calculate_ratio(self, context):
        """Determines the width of the widget from the event data."""
        if self._parent:
            self._parent.calculate_ratio(context)
        else:
            if not self.ratio:
                # Find the higest ratio of rendered string length to time
                if self.activities:
                    for activity in self.activities:
                        time = get_decimal_hours(activity.end.time - activity.start.time)
                        extents = context.text_extents(activity.label)
                        ratio = (extents[2] + 2 * self.x_pad) / time
                        if ratio > self.ratio:
                            self.ratio = ratio
                        if self.height < extents[3] + 2 * self.y_pad:
                            self.height = extents[3] + 2 * self.y_pad
                self.ratio *= self.zoom

    def calculate_heights(self):
        """Determines the height of each activity based on whether it overlaps others."""
        group = []
        max = 0
        activities = {}
        for activity in self.activities:
            activity.count = 1
             
        events = self.markers.values()
        events.sort(key=operator.attrgetter('time'))

        for event in events:
            for prior in event.prior:
                if prior.y in activities:
                    del activities[prior.y]
            if len(activities) == 0:
                for activity in group:
                    if activity.count < max:
                        activity.count = max
                group = []
                max = 0
            for subsequent in event.subsequent:
                group.append(subsequent)
                for index in range(len(activities) + 1):
                    if index not in activities:
                        activities[index] = subsequent
                        subsequent.y = index
                        break
            if len(activities) > max:
                max = len(activities)

    def change_start(self, activity, new_start):
        """Changes the start event for an activity."""
        old_start = activity.start
        activity.start = new_start
        old_start.subsequent.remove(activity)
        new_start.subsequent.add(activity)
        if activity.the_activity:
            activity.the_activity.change_start(new_start.the_event)
        self.delete_event(old_start)

    def change_end(self, activity, new_end):
        """Changes the end event for an activity."""
        old_end = activity.end
        activity.end = new_end
        old_end.prior.remove(activity)
        new_end.prior.add(activity)
        if activity.the_activity:
            activity.the_activity.change_end(new_end.the_event)
        self.delete_event(old_end)

    def new_marker(self, event, label=None):
        """Adds an event to the timeline."""
        marker = Timeline.Marker(event, label)
        self.markers[event] = marker
        self.update_start_end()
        return marker

    def get_event(self, event, create=True):
        """Finds an event on the timeline based on the external event."""
        marker = self.markers.get(event)
        if (not marker) and create:
            marker = self.new_marker(event)
        return marker

    def delete_event(self, event):
        """Deletes an event from the timeline."""
        if not event.prior and not event.subsequent:
            del self.markers[event.the_event]
            self.update_start_end()

    def update_activity_list(self, activities):
        """Updates the internal activities based on the given list."""
        # If the date has changed delete all activities
        if self.start.date() != self.date:
            self.activities = []
            self.markers = {}

        # Remove any internal activities that have an invalid reference to an
        # external activity
        #
        for internal_activity in self.activities:
            activity = internal_activity.the_activity
            if activity:
                if activity in activities:
                    start = self.get_event(activity.start_event)
                    if activity.end_event:
                        end = self.get_event(activity.end_event)
                    else:
                        end = None
                        
                    if internal_activity.start != start:
                        old_start = internal_activity.start
                        internal_activity.start = start
                        start.subsequent.add(internal_activity)
                        old_start.subsequent.remove(internal_activity)
                        self.delete_event(old_start)
                    if internal_activity.end != end:
                        old_end = internal_activity.end
                        internal_activity.end = end
                        end.prior.add(internal_activity)
                        old_end.prior.remove(internal_activity)
                        self.delete_event(old_end)
                else:
                    internal_activity.the_activity = None
                    internal_activity.delete()

        # Add any external activities without a corresponding internal activity
        #
        referenced = [activity.the_activity for activity in self.activities
                      if activity.the_activity]
        start = datetime.datetime.combine(self.date, datetime.time())
        end = start + datetime.timedelta(1)

        for external_activity in activities:
            if (external_activity.start_event.time > end or
                external_activity.end_event.time < start):
                continue
            if external_activity not in referenced:
                self.new_activity(external_activity,
                                  external_activity.start_event,
                                  external_activity.end_event)                

    def new_activity(self, activity, start_event, end_event, label=None):
        """Adds an event to the timeline."""
        start = self.get_event(start_event)
        if end_event:
            end = self.get_event(end_event)
        else:
            end = self.new_marker(None)
            if activity and not activity.is_active():
                Message.show('Invalid activity',
                             'Activity %d (%s) is not the current activiy but does not have an end event!' % (activity.get_id(), label),
                             gtk.MESSAGE_WARNING)
        if not label:
            if activity:
                label = activity.task.name
            else:
                label = 'New Activity'

        new_activity = None
        for internal_activity in self.activities:
            if internal_activity.the_activity == activity:
                new_activity = activity
                break
        if new_activity:
            new_activity.label = label
        else:
            new_activity = Timeline.Activity(activity, start, end, label, self.session)
        self.activities.append(new_activity)
        self.update_start_end()
        return new_activity

class Week_Selector(gtk.HBox):

    def __init__(self, date=None):
        """Connect signals for the widget."""
        gtk.HBox.__init__(self, False, 0)
        icon = gtk.Image()
        pixbuf = gtk.gdk.pixbuf_new_from_file(config.get('image_path_16') + "go-jump-locationbar.png")
        icon.set_from_pixbuf(pixbuf)
        self.go_button = gtk.Button()
        self.go_button.set_image(icon)
        self.go_button.set_tooltip_text('Display timesheet for selected week')
        self.go_button.set_relief(gtk.RELIEF_NONE)
        self.go_button.set_focus_on_click(False)
        icon = gtk.Image()
        pixbuf = gtk.gdk.pixbuf_new_from_file(config.get('image_path_16') + "go-jump-today.png")
        icon.set_from_pixbuf(pixbuf)
        self.today_button = gtk.Button()
        self.today_button.set_image(icon)
        self.today_button.set_tooltip_text('Jump to current week')
        self.today_button.set_relief(gtk.RELIEF_NONE)
        self.today_button.set_focus_on_click(False)
        icon = gtk.Image()
        pixbuf = gtk.gdk.pixbuf_new_from_file(config.get('image_path_16') + "go-previous.png")
        icon.set_from_pixbuf(pixbuf)
        self.previous_button = gtk.Button()
        self.previous_button.set_image(icon)
        self.previous_button.set_tooltip_text('Select previous week')
        self.previous_button.set_relief(gtk.RELIEF_NONE)
        self.previous_button.set_focus_on_click(False)
        icon = gtk.Image()
        pixbuf = gtk.gdk.pixbuf_new_from_file(config.get('image_path_16') + "go-next.png")
        icon.set_from_pixbuf(pixbuf)
        self.next_button = gtk.Button()
        self.next_button.set_image(icon)
        self.next_button.set_tooltip_text('Select following week')
        self.next_button.set_relief(gtk.RELIEF_NONE)
        self.next_button.set_focus_on_click(False)

        self.calendar = gtk.Calendar()
        self.calendar.set_display_options(gtk.CALENDAR_SHOW_HEADING
                                          | gtk.CALENDAR_SHOW_DAY_NAMES
                                          | gtk.CALENDAR_SHOW_WEEK_NUMBERS
                                          | gtk.CALENDAR_WEEK_START_MONDAY)
        self.entry = gtk.Entry()

        image = gtk.Image()
        pixbuf = gtk.gdk.pixbuf_new_from_file(config.get('image_path_16') + "view-calendar-week.png")
        image.set_from_pixbuf(pixbuf)
        self.button = gtk.Button()
        self.button.set_image(image)
        self.button.set_tooltip_text('Pick date')
        self.button.set_relief(gtk.RELIEF_NONE)
        #self.button.set_focus_on_click(False)
        icon = gtk.Image()
        self.calendar_window = gtk.Window(gtk.WINDOW_TOPLEVEL)
        self.display = False

        self.calendar_window.set_position(gtk.WIN_POS_MOUSE)
        self.calendar_window.set_decorated(False)
        self.calendar_window.add(self.calendar)
        self.calendar_window.connect('focus-out-event', self.focus_out_event)

        self.entry.set_width_chars(10)

        label = gtk.Label("Week Ending")
        self.pack_start(label, False, False, 0)
        self.pack_start(self.entry, True, True, 0)
        self.pack_start(self.previous_button, False, False, 0)
        self.pack_start(self.next_button, False, False, 0)
        self.pack_start(self.button, False, False, 0)
        self.pack_start(self.today_button, False, False, 0)
        self.pack_start(self.go_button, False, False, 0)

        self.__connect_signals()
        self.set_date(date)
        self.update_entry()

    def __connect_signals(self):
        """Connect signals for the widget."""
        self.previous_button.connect('clicked', self.change_week_cb, '-')
        self.go_button.connect('clicked', self.change_week_cb)
        self.today_button.connect('clicked', self.change_week_cb, 'today')
        self.next_button.connect('clicked', self.change_week_cb, '+')
        self.day_selected_handle = self.calendar.connect('day-selected', self.update_entry)
        self.day_selected_double_handle = self.calendar.connect('day-selected-double-click', self.hide_widget)
        self.clicked_handle = self.button.connect('clicked', self.show_widget)
        self.activate = self.entry.connect('activate', self.update_calendar)
        self.focus_out = self.entry.connect('focus-out-event', self.focus_out_event)

    def __block_signals(self):
        """Connect signals for the widget."""
        self.calendar.handler_block(self.day_selected_handle)
        self.calendar.handler_block(self.day_selected_double_handle)
        self.button.handler_block(self.clicked_handle)
        self.entry.handler_block(self.activate)
        self.entry.handler_block(self.focus_out)

    def __unblock_signals(self):
        """Connect signals for the widget."""
        self.calendar.handler_unblock(self.day_selected_handle)
        self.calendar.handler_unblock(self.day_selected_double_handle)
        self.button.handler_unblock(self.clicked_handle)
        self.entry.handler_unblock(self.activate)
        self.entry.handler_unblock(self.focus_out)

    def change_week_cb(self, widget, data = None):
        """Changes the week based on the data."""
        year, month, day = self.calendar.get_date()
        date = datetime.date(year, month + 1, day)
        if data == '+':
            date += datetime.timedelta(7)
        elif data == '-':
            date -= datetime.timedelta(7)
        elif data == 'today':
            delta = 6 - datetime.datetime.today().weekday()
            date = datetime.datetime.today()
            date += datetime.timedelta(delta)
        if data:
            self.set_date(date)
            self.update_entry()

    def get_text(self):
        """Returns the content of the entry field."""
        return self.entry.get_text()

    def set_date(self, date):
        """Sets the date associated with the widget."""
        if not date:
            date = datetime.date.today()
        delta = 6 - date.weekday()
        date += datetime.timedelta(delta)
        self.__block_signals()
        self.calendar.select_day(1)
        self.calendar.select_month(date.month - 1, date.year)
        self.calendar.select_day(date.day)
        self.__unblock_signals()

    def get_date(self):
        """Returns the currently selected week."""
        year, month, day = self.calendar.get_date()
        date = datetime.date(year, month + 1, day)
        return date

    def hide_widget(self, *args):
        self.calendar_window.hide_all()

    def show_widget(self, *args):
        self.calendar_window.show_all()

    def update_entry(self, *args):
        """Displays the selected date."""
        year, month, day = self.calendar.get_date()
        date = datetime.date(year, month + 1, day)
        weekday = 6 - date.weekday()
        date += datetime.timedelta(weekday)
        date_format = date.strftime(config.get('date_format'))
        self.entry.set_text(date_format)
        self.__block_signals()
        self.calendar.select_day(1)
        self.calendar.select_month(date.month - 1, date.year)
        self.calendar.select_day(date.day)
        self.__unblock_signals()

    def update_calendar(self, *args):
        """Update the widget's date from the entry field."""
        text = self.entry.get_text()
        date = string_to_date(text)
        self.set_date(date)
        self.update_entry()

    def focus_out_event(self, widget, event):
        """Hide calendar when focus is lost."""
        self.update_calendar()
        self.hide_widget()

class Day_Selector(gtk.HBox):

    def __init__(self, date=None):
        """Connect signals for the widget."""
        gtk.HBox.__init__(self, False, 0)
        icon = gtk.Image()
        pixbuf = gtk.gdk.pixbuf_new_from_file(config.get('image_path_16') + "go-jump-today.png")
        icon.set_from_pixbuf(pixbuf)
        self.today_button = gtk.Button()
        self.today_button.set_image(icon)
        self.today_button.set_tooltip_text("Insert today's date")
        self.today_button.set_relief(gtk.RELIEF_NONE)
        self.today_button.set_focus_on_click(False)

        self.calendar = gtk.Calendar()
        self.calendar.set_display_options(gtk.CALENDAR_SHOW_HEADING
                                          | gtk.CALENDAR_SHOW_DAY_NAMES
                                          | gtk.CALENDAR_SHOW_WEEK_NUMBERS
                                          | gtk.CALENDAR_WEEK_START_MONDAY)
        self.entry = gtk.Entry()

        image = gtk.Image()
        pixbuf = gtk.gdk.pixbuf_new_from_file(config.get('image_path_16') + "view-calendar-week.png")
        image.set_from_pixbuf(pixbuf)
        self.button = gtk.Button()
        self.button.set_image(image)
        self.button.set_tooltip_text('Pick on calendar')
        self.button.set_relief(gtk.RELIEF_NONE)
        #self.button.set_focus_on_click(False)
        icon = gtk.Image()
        self.calendar_window = gtk.Window(gtk.WINDOW_TOPLEVEL)
        self.display = False

        self.calendar_window.set_position(gtk.WIN_POS_MOUSE)
        self.calendar_window.set_decorated(False)
        self.calendar_window.add(self.calendar)
        self.calendar_window.connect('focus-out-event', self.focus_out_event)

        self.entry.set_width_chars(10)

        self.pack_start(self.entry, True, True, 0)
        self.pack_start(self.button, False, False, 0)
        self.pack_start(self.today_button, False, False, 0)

        self.__connect_signals()
        self.set_date(date)
        self.update_entry()

    def __connect_signals(self):
        """Connect signals for the widget."""
        self.today_button.connect('clicked', self.change_date_cb, 'today')
        self.day_selected_handle = self.calendar.connect('day-selected', self.update_entry)
        self.day_selected_double_handle = self.calendar.connect('day-selected-double-click', self.hide_widget)
        self.clicked_handle = self.button.connect('clicked', self.show_widget)
        self.activate = self.entry.connect('activate', self.update_calendar)
        self.focus_out = self.entry.connect('focus-out-event', self.focus_out_event)

    def __block_signals(self):
        """Connect signals for the widget."""
        self.calendar.handler_block(self.day_selected_handle)
        self.calendar.handler_block(self.day_selected_double_handle)
        self.button.handler_block(self.clicked_handle)
        self.entry.handler_block(self.activate)
        self.entry.handler_block(self.focus_out)

    def __unblock_signals(self):
        """Connect signals for the widget."""
        self.calendar.handler_unblock(self.day_selected_handle)
        self.calendar.handler_unblock(self.day_selected_double_handle)
        self.button.handler_unblock(self.clicked_handle)
        self.entry.handler_unblock(self.activate)
        self.entry.handler_unblock(self.focus_out)

    def change_date_cb(self, widget, data = None):
        """Changes the week based on the data."""
        year, month, day = self.calendar.get_date()
        date = datetime.date(year, month + 1, day)
        if data == '+':
            date += datetime.timedelta(1)
        elif data == '-':
            date -= datetime.timedelta(1)
        elif data == 'today':
            date = datetime.datetime.today()
        if data:
            self.set_date(date)
            self.update_entry()

    def get_text(self):
        """Returns the content of the entry field."""
        return self.entry.get_text()

    def set_date(self, date):
        """Sets the date associated with the widget."""
        self.__block_signals()
        self.date = date
        if date == None:
            self.entry.set_text('Not set')
        else:
            self.calendar.select_day(1)
            self.calendar.select_month(date.month - 1, date.year)
            self.calendar.select_day(date.day)
            self.update_entry()
        self.__unblock_signals()

    def get_date(self):
        """Returns the currently selected week."""
        year, month, day = self.calendar.get_date()
        date = datetime.datetime(year, month + 1, day)
        return date

    def hide_widget(self, *args):
        """Hides calendar widget."""
        self.calendar_window.hide_all()

    def show_widget(self, *args):
        """Displays calendar widget."""
        self.calendar_window.show_all()

    def update_entry(self, *args):
        """Displays the selected date."""
        if not self.date:
            self.entry.set_text('Not set')
        else:
            year, month, day = self.calendar.get_date()
            date = datetime.date(year, month + 1, day)
            date_format = date.strftime(config.get('date_format'))
            self.entry.set_text(date_format)
            self.__block_signals()
            self.calendar.select_day(1)
            self.calendar.select_month(date.month - 1, date.year)
            self.calendar.select_day(date.day)
            self.__unblock_signals()

    def update_calendar(self, *args):
        """Update the widget's date from the entry field."""
        text = self.entry.get_text()
        date = string_to_date(text)
        self.set_date(date)
        self.update_entry()
        #self.hide_widget()

    def focus_out_event(self, widget, event):
        """Hide calendar when focus is lost."""
        self.update_calendar()
        self.hide_widget()

class Notes (gtk.ScrolledWindow, dragndrop.Drop_Target_Interface):
    """A scrolled text widget."""

    def __init__ (self, buffer=None):
        gtk.ScrolledWindow.__init__(self)
        self.set_shadow_type(gtk.SHADOW_IN)
        self.set_policy(gtk.POLICY_AUTOMATIC, gtk.POLICY_AUTOMATIC)
        self.view = gtk.TextView(buffer)
        self.view.set_wrap_mode(gtk.WRAP_WORD_CHAR)
        self.connect('key-release-event', self.key_release_cb)
        self.view.set_property('has-tooltip', True)
        self.view.connect('query-tooltip', self.query_tooltip_cb)
        self.view.connect('populate-popup', self.populate_popup_cb)
        text_buffer = self.view.get_buffer()
        self.insert_handler = text_buffer.connect_after('insert-text',
                                                        self.insert_text_cb)
        self.user_tag = text_buffer.create_tag(None, foreground='Blue')

        self.show_markup = text_buffer.create_tag('show_markup')
        self.markup = text_buffer.create_tag('markup')

        self.tags = {'edit':None, 'link':None}
        self.style_tags = {}
        self.user_tags = set()
        self.show_tags = False

        self.add_drop_format(dragndrop.TEXT, self.drop_text)
        self.add_drop_format(dragndrop.FILES, self.drop_files)
        self.add(self.view)

    def populate_popup_cb(self, view, menu):
        """Add custom items to view popup menu."""
        show_tags = gtk.CheckMenuItem('Show Tags')
        show_tags.set_active(self.show_tags)
        show_tags.connect('toggled', self.show_tags_cb)
        menu.append(show_tags)
        show_tags.show_all()
        buffer = self.view.get_buffer()
        selection = buffer.get_selection_bounds()
        cursor = buffer.get_iter_at_mark(buffer.get_insert())
        tags = set([tag for tag in cursor.get_tags() if tag.get_property('name')])
        tags = tags - set(self.user_tags)
        if selection or tags:
            tag_menu = gtk.Menu()
            menu_item = gtk.MenuItem('Tags')
            menu_item.set_submenu(tag_menu)
            
            remove_all = gtk.MenuItem('Remove All')
            remove_all.connect('activate', self.remove_tags_cb, selection)
            tag_menu.append(remove_all)

            if selection or tags:
                tag_menu.append(gtk.SeparatorMenuItem())

            if selection and self.style_tags:
                add_menu = gtk.Menu()
                add = gtk.MenuItem('Add')
                add.set_submenu(add_menu)
                tag_menu.append(add)
            
                for tag in self.style_tags.keys():
                    apply_tag = gtk.MenuItem(tag[:-1])
                    apply_tag.connect('activate', self.apply_tag_cb, tag, selection)
                    add_menu.append(apply_tag)
            
            if tags:
                remove_menu = gtk.Menu()
                remove = gtk.MenuItem('Remove')
                remove.set_submenu(remove_menu)
                tag_menu.append(remove)
            
                for tag in tags:
                    remove_tag = gtk.MenuItem(tag.get_property('name')[:-1])
                    remove_tag.connect('activate', self.remove_tag_cb, tag, selection)
                    remove_menu.append(remove_tag)
            
            menu.append(menu_item)
            menu_item.show_all()
            
    def apply_tag_cb(self, item, tag, selection):
        """Adds a tag to the specified range."""
        buffer = self.view.get_buffer()
        tag = buffer.get_tag_table().lookup(tag)
        if tag:
            name = tag.get_property('name')
            if name in self.style_tags:
                name += ','.join(['%s=%s' % (v,k) for (v,k) in self.style_tags[name].iteritems()])
            if selection[0].has_tag(self.show_markup):
                mark = buffer.create_mark(None, selection[1]) 
                self.view.get_buffer().insert(selection[0], '[%s]' % (name))
                selection = (selection[0], buffer.get_iter_at_mark(mark))
                buffer.delete_mark(mark)
            self.view.get_buffer().apply_tag(tag, *selection)
            name = name[:name.find(':') + 1]
            if selection[0].has_tag(self.show_markup):
                mark = buffer.create_mark(None, selection[0]) 
                self.view.get_buffer().insert(selection[1], '[%s]' % (name))
                selection = (buffer.get_iter_at_mark(mark), selection[1])
                buffer.delete_mark(mark)

    def remove_tag_cb(self, item, tag, selection):
        """Adds a tag to the specified range."""
        buffer = self.view.get_buffer()
        if selection:
            start, end = selection
        else:
            start = buffer.get_iter_at_mark(buffer.get_insert())
            end = buffer.get_iter_at_mark(buffer.get_insert())
        
            if start.has_tag(tag):
                start.backward_to_tag_toggle(tag)
                
            if end.has_tag(tag):
                end.forward_to_tag_toggle(tag)

        buffer.remove_tag(tag, start, end)


    def remove_tags_cb(self, item, selection):
        """Removes all the custom tags on the selected range."""
        buffer = self.view.get_buffer()
        for tag_name in self.style_tags.keys():
            tag = buffer.get_tag_table().lookup(tag_name)
            self.view.get_buffer().remove_tag(tag, *selection)

    def show_tags_cb(self, item):
        """Update the state of the show markup flag."""
        self.show_tags = item.get_active()
        the_buffer = self.view.get_buffer()
        bounds = the_buffer.get_bounds()
        if self.show_tags:
            the_buffer.apply_tag(self.show_markup, *bounds)
            self.display_tags(the_buffer, bounds)
        else:
            the_buffer.remove_tag(self.show_markup, *bounds)
            self.process_tags(the_buffer, bounds)
            self.hide_tags(the_buffer, bounds)

    def insert_text_cb(self, text_buffer, position, text, length):
        """Adds a tag for the current user and date to text added to the buffer."""
        text_buffer.handler_block(self.insert_handler)
        user_tag = user.get_id()[0] + ' ' + str(datetime.date.today())
        tag = text_buffer.get_tag_table().lookup('edit:' + user_tag)
        if not tag:
            tag = text_buffer.create_tag('edit:' + user_tag)
            self.user_tags.add(tag)
        start = position.copy()
        start.backward_chars(length)
        text_buffer.apply_tag(tag, start, position)
        text_buffer.handler_unblock(self.insert_handler)

    def key_release_cb(self, widget, event):
        """Perform custom key processing."""
        if (event.state & gtk.gdk.CONTROL_MASK and
            event.keyval == gtk.keysyms.l):
            self.selection_to_link()
        elif event.keyval == gtk.keysyms.F9:
            if event.state & gtk.gdk.CONTROL_MASK:
                # Force showing of tags
                self.change_show_tags(True)
            elif event.state & gtk.gdk.SHIFT_MASK:
                # Force hiding of tags
                self.change_show_tags(False)
            else:
                # Toggle show tag setting
                self.change_show_tags()            
            return True
        elif event.keyval == gtk.keysyms.F10:
            if event.state & gtk.gdk.CONTROL_MASK:
                self.rationalise_tags()
            else:
                self.evaluate_tag()

    def rationalise_tags(self):
        """Remove unused formatting tags."""
        buffer = self.view.get_buffer()
        tag_table = buffer.get_tag_table()
        self.tag_list = []
        tag_table.foreach(lambda tag, data: self.tag_list.append(tag))
        tag_list = [tag for tag in self.tag_list if tag.get_property('name')]
        del self.tag_list
        for tag in tag_list:
            name = tag.get_property('name')
            if name in self.style_tags:
                start = buffer.get_start_iter()
                if not start.forward_to_tag_toggle(tag):
                    tag_table.remove(tag)
                    del self.style_tags[name]  
 
    def evaluate_tag(self):
        """Set the tag fomatting to the specified values at the cursor."""
        buffer = self.view.get_buffer()
        text = buffer.get_text(*buffer.get_bounds())
        offset = buffer.get_iter_at_mark(buffer.get_insert()).get_offset()
        start = text[:offset].rfind('[') + 1
        end = offset + text[offset:].find(']')
        if start > 0 and end >= offset:
            tag_text = text[start:end]
            pos = tag_text.find(':')
            if pos > 0:
                tag_name = tag_text[:pos + 1]
                properties = tag_text[pos + 1:]
                self.style_tags[tag_name] = {}
                tag = buffer.get_tag_table().lookup(tag_name)
                if tag:
                    # Reset existing properties
                    tag.set_properties(background_full_height_set = False,
                                       background_set = False,
                                       background_stipple_set = False,
                                       editable_set = False,
                                       family_set = False,
                                       foreground_set = False,
                                       foreground_stipple_set = False,
                                       indent_set = False,
                                       invisible_set = False,
                                       justification_set = False,
                                       language_set = False,
                                       left_margin_set = False,
                                       paragraph_background_set = False,
                                       pixels_above_lines_set = False,
                                       pixels_below_lines_set = False,
                                       pixels_inside_wrap_set = False,
                                       right_margin_set = False,
                                       rise_set = False,
                                       scale_set = False,
                                       size_set = False,
                                       stretch_set = False,
                                       strikethrough_set = False,
                                       style_set = False,
                                       tabs_set = False,
                                       underline_set = False,
                                       variant_set = False,
                                       weight_set = False,
                                       wrap_mode_set = False)
                else:
                    tag = buffer.create_tag(tag_name)
                self.set_tag_properties(tag, properties)
 
    def change_show_tags(self, value=None):
        """Changes the value of the show tags variable and updates the display."""
        buffer = self.view.get_buffer()
        bounds = buffer.get_selection_bounds()
        tag_bounds = self.style_tags.keys() + ['show_markup', 'markup']
        if not bounds:
            cursor = buffer.get_iter_at_mark(buffer.get_insert())
            bounds = (cursor, cursor)
        start_iter, end_iter = bounds
        tags = [tag for tag in start_iter.get_tags()
                if tag.get_property('name') in tag_bounds]
        current = start_iter.copy()
        for tag in tags:
            start = current.copy()
            if not start.begins_tag(tag):
                start.backward_to_tag_toggle(tag)
                if start.compare(start_iter) == -1:
                    start_iter = start
        
        tags = [tag for tag in end_iter.get_tags()
                if tag.get_property('name') in tag_bounds]
        current = end_iter.copy()
        for tag in tags:
            end = current.copy()
            if not end.begins_tag(tag):
                end.forward_to_tag_toggle(tag)
                if end.compare(end_iter) == 1:
                    end_iter = end
        bounds = (start_iter, end_iter)
        if value != None:
            show_tags = value
        else:
            cursor = buffer.get_iter_at_mark(buffer.get_insert())
            show_tags = not cursor.has_tag(self.show_markup)

        if show_tags:
            buffer.apply_tag(self.show_markup, *bounds)
            self.display_tags(buffer, bounds)
        else:
            buffer.remove_tag(self.show_markup, *bounds)
            self.process_tags(buffer, bounds)
            self.hide_tags(buffer, bounds)

    def display_tags(self, text_buffer, bounds):
        """Insert the string replesentation of tags into the buffer."""
        # Hide tags already displayed
        start_mark = text_buffer.create_mark(None, bounds[0])
        end_mark = text_buffer.create_mark(None, bounds[1])
        
        #self.hide_tags(text_buffer, bounds)
        style_tags = self.style_tags.keys()
        start = text_buffer.get_iter_at_mark(start_mark)
        tags = start.get_toggled_tags(True)
        tags = [tag for tag in tags if tag.get_property('name') in style_tags]
        known_tags = []
        for tag in tags:
            text = '[%s]' % (self.tag_to_string(tag, known_tags))
            known_tags.append(tag)
            self.add_text(text, start, [None, self.markup, self.show_markup])
        start.forward_to_tag_toggle(None)
        while start.compare(text_buffer.get_iter_at_mark(end_mark)) < 0:
            tags = start.get_toggled_tags(True)
            tags = [tag for tag in tags if tag.get_property('name') in style_tags]
            for tag in tags:
                text = '[%s]' % (self.tag_to_string(tag, known_tags))
                known_tags.append(tag)
                self.add_text(text, start, [None, self.markup, self.show_markup])
            tags = start.get_toggled_tags(False)
            tags = [tag for tag in tags if tag.get_property('name') in style_tags]
            for tag in tags:
                text = '[%s]' % (tag.get_property('name'))
                self.add_text(text, start, [None, self.markup, self.show_markup])
            start.forward_to_tag_toggle(None)

        tags = start.get_toggled_tags(False)
        tags = [tag for tag in tags if tag.get_property('name') in style_tags]
        for tag in tags:
            text = '[%s]' % (tag.get_property('name'))
            self.add_text(text, start, [None, self.markup, self.show_markup])
        text_buffer.delete_mark(start_mark)
        text_buffer.delete_mark(end_mark)

    def process_tags(self, text_buffer, bounds):
        """Process any style tags within the specified buffer bounds."""
        text = text_buffer.get_text(*bounds)
        tag_table = text_buffer.get_tag_table()
        blocks = []
        close_brace = 0
        open_brace = text.find('[')
        while open_brace != -1:
            count = 1
            while count:
                close_brace = text.find(']', open_brace)
                string = text[open_brace + 1:close_brace]
                count = string.count('[') - string.count(']')
            if string:
                blocks.append(string)
            close_brace += 1
            open_brace = text.find('[', close_brace)
        string = text[close_brace:]
        if string:
            blocks.append(string)

        for content in blocks:
            if content.startswith('link:'):
                self.set_link_tag(content[5:])
            elif content.startswith('edit:'):
                tag_type = content[:content.find(':')]
                tag_name = content[content.find(':') + 1:]
                self.set_tag(tag_name, tag_type)
            else:
                pos = content.find(':')
                if pos != -1:
                    tag_name = content[:pos + 1]
                    tag_properties = content[pos + 1:]
                    if tag_properties:
                        tag = tag_table.lookup(tag_name)
                        if not tag:
                            tag = text_buffer.create_tag(tag_name)
                            if tag_name not in self.tags:
                                self.style_tags[tag_name] = {}
                        self.set_tag_properties(tag, tag_properties)
                        self.tags[tag_name] = tag

    def hide_tags(self, text_buffer, bounds):
        """Remove any style tags from display."""
        start = bounds[0]
        mark = text_buffer.create_mark(None, bounds[1])
        if not start.has_tag(self.markup):
            start.forward_to_tag_toggle(self.markup)
        while start.compare(text_buffer.get_iter_at_mark(mark)) < 0:
            end = start.copy()
            end.forward_to_tag_toggle(self.markup)
            text_buffer.delete(start, end)
            if not start.has_tag(self.markup):
                start.forward_to_tag_toggle(self.markup)
        text_buffer.delete_mark(mark)

    def set_tag(self, name, tag_type='edit'):
        """Return a tag of the given name, creating it if necessary."""
        if name:
            text_buffer = self.view.get_buffer()
            tag = text_buffer.get_tag_table().lookup('edit:' + name)
            if not tag:
                tag = text_buffer.create_tag('edit:' + name)
            self.user_tags.add(tag)
        else:
            tag = None
        self.tags[tag_type] = tag

    def set_link_tag(self, link):
        """Creates a tag for a link within the text."""
        if link:
            text_buffer = self.view.get_buffer()
            new_tag = text_buffer.get_tag_table().lookup('link:' + link)
            if not new_tag:
                new_tag = text_buffer.create_tag('link:' + link)
                new_tag.link = link
                new_tag.connect('event', self.tag_cb)
                new_tag.set_property('underline', True)
                new_tag.set_property('foreground', 'Blue')
        else:
            new_tag = None
        self.tags['link'] = new_tag

    def add_text(self, text, position=None, tags=[]):
        """Add a markup block of text."""
        text_buffer = self.view.get_buffer()
        text_buffer.handler_block(self.insert_handler)
        if not position:
            position = text_buffer.get_end_iter()
        if tags and tags[0] == None:
            tags = tags[1:]
        else:
            tags = [tag for tag in self.tags.values() if tag] + tags
        if tags:
            text_buffer.insert_with_tags(position, text, *tags)
        else:
            text_buffer.insert(position, text)
        text_buffer.handler_unblock(self.insert_handler)

    def get_handle(self):
        """Returns the window manager handle for this window."""
        return self.window.handle

    def drop_files(self, content, point):
        """Adds links to the listed files to the buffer."""
        files = dragndrop.get_files(content)
        buffer = self.view.get_buffer()
        location = gtk.gdk.window_at_pointer()
        position = self.view.get_iter_at_location(location[1], location[2])
        if not position:
            position = buffer.get_end_iter()

        for file_name in files:
            self.set_link_tag(file_name)
            base_name = os.path.basename(file_name)
            edit_tag = user.get_id()[0] + ' ' + str(datetime.date.today())
            self.set_tag(edit_tag)
            self.add_text(base_name, position)
            self.set_link_tag(None)

    def tag_to_string(self, tag, known_tags=[]):
        """Returns a string representation of the tag formatting."""
        tag_name = tag.get_property('name')
        if tag_name:
            if tag_name in self.style_tags:
                if tag not in known_tags:
                    tag_name += ','.join(['%s=%s' % (v,k) for (v,k) in self.style_tags[tag_name].iteritems()])
                else:
                    tag_name += 'Start'
        return tag_name

    def parse_buffer(self, bounds=None):
        """Returns the buffer text converting tags to markup."""
        buffer = self.view.get_buffer()
        defined_tags = set()
        
        if bounds:
            text_iter = bounds[0].copy()
            end_iter = bounds[1]
        else:
            text_iter, end_iter = buffer.get_bounds()

        if text_iter.has_tag(self.markup):
            text_iter.forward_to_tag_toggle(self.markup)
        start_iter = text_iter.copy()
        text = ''
        known_tags = []
        while text_iter.compare(end_iter) == -1 and not text_iter.is_end():
            text_iter.forward_to_tag_toggle(None)
            if text_iter.has_tag(self.markup):
                text_iter.forward_to_tag_toggle(self.markup)
                start = text_iter.copy()
                text_iter.forward_to_tag_toggle(None)
            if text_iter.has_tag(self.show_markup):
                text_iter.forward_to_tag_toggle(self.show_markup)
                start = text_iter.copy()
                text_iter.forward_to_tag_toggle(None)

            string = buffer.get_text(start_iter, text_iter)            
            tags = set(start_iter.get_toggled_tags(toggled_on=True))
            start_tags = [self.tag_to_string(tag, known_tags) for tag in tags
                          if tag.get_property('name')]
            known_tags += tags
            end_tags = text_iter.get_toggled_tags(toggled_on=False)
            end_tags = [tag.get_property('name') for tag in end_tags
                        if tag.get_property('name')]

            intro = '[' + ']['.join(start_tags) + ']' if start_tags else ''
            tail = [tag[:tag.find(':') + 1] for tag in end_tags if tag.find(':') > 0]
            tail = '[' + ']['.join(tail) + ']' if tail else ''
            text += intro + string + tail
            start_iter = text_iter.copy()
        return text

    def set_tag_properties(self, tag, properties):
        """Converts the properties sting into the appropriate types and apply them to tag."""
        tag_name = tag.get_property('name')
        if '=' in properties:
            for name, value in [(tp.split('=')[0], tp.split('=')[1])
                                for tp in properties.split(',')]:
                self.style_tags[tag_name][name] = value
                try:
                    value = tag.get_property(name).__class__(value)
                except (TypeError, ValueError):
                    try:
                        value = eval(value)
                    except NameError:
                        # Assume string type
                        pass
                try:
                    tag.set_property(name, value)
                except TypeError:
                    Message.Show('Tag Property Error',
                                 'Failed to interpret poperty %s for tag %s!'
                                  % (name, str(value)),
                                 gtk.MESSAGE_WARNING)

    def display_notes(self, text, bounds=None, show_tags=None):
        """Display the task notes adding tasks for file links."""
        if not show_tags:
            show_tags = self.show_tags
        if show_tags:
            self.tags['markup'] = self.show_markup
        position = bounds[0] if bounds else None
        buffer = self.view.get_buffer()
        tag_table = buffer.get_tag_table()
        blocks = []
        close_brace = 0
        open_brace = text.find('[')
        while open_brace != -1:
            string = text[close_brace:open_brace]
            if string:
                blocks.append((string, True))
            count = 1
            while count:
                close_brace = text.find(']', open_brace)
                string = text[open_brace + 1:close_brace]
                count = string.count('[') - string.count(']')
            if string:
                blocks.append((string, False))
            close_brace += 1
            open_brace = text.find('[', close_brace)
        string = text[close_brace:]
        if string:
            blocks.append((string, False))

        for block in blocks:
            if not block:
               continue
            content = block[0]
            is_text = block[1]
            if is_text:
                self.add_text(content, position)
            elif content.startswith('link:'):
                if show_tags:
                    self.add_text('[%s]' % (content), position, [None, self.show_markup])
                self.set_link_tag(content[5:])
            elif content.startswith('edit:'):
                if show_tags:
                    self.add_text('[%s]' % (content), position, [None, self.show_markup])
                tag_type = content[:content.find(':')]
                tag_name = content[content.find(':') + 1:]
                self.set_tag(tag_name, tag_type)
            else:
                pos = content.find(':')
                if pos == -1:
                    self.add_text('[%s]' % (content), position)
                else:
                    tag_name = content[:pos + 1]
                    tag_properties = content[pos + 1:]
                    if tag_properties:
                        if show_tags:
                            self.add_text('[%s]' % (content), position, [None, self.show_markup])
                        tag = tag_table.lookup(tag_name)
                        if not tag:
                            tag = buffer.create_tag(tag_name)
                            if tag_name not in self.tags:
                                self.style_tags[tag_name] = {}
                        self.set_tag_properties(tag, tag_properties)
                        self.tags[tag_name] = tag
                    else:
                        if tag_name in self.style_tags:
                            if tag_name in self.tags:
                                del self.tags[tag_name]
                                if show_tags:
                                    self.add_text('[%s]' % (content), position, [None, self.show_markup])
                            else:
                                self.add_text('[%s]' % (content), position)
                                Message.show('Parse Error',
                                             'Unmatched close tag %s' % (tag_name),
                                             gtk.MESSAGE_WARNING) 
                        else:
                            self.add_text('[%s]' % (content), position)
                            Message.show('Parse Error',
                                         'Unknown tag %s' % (tag_name),
                                         gtk.MESSAGE_WARNING) 
        if 'markup' in self.tags:
            del self.tags['markup']


    def selection_to_link(self):
        """Attempts to convert the sellected text into a link."""
        buffer = self.view.get_buffer()
        bounds = buffer.get_selection_bounds()
        if bounds:
            selection = buffer.get_text(bounds[0], bounds[1])
            selection = selection.strip()
            if os.path.exists(selection):
                name = os.path.basename(selection)
            else:
                name = selection
            buffer.delete(bounds[0], bounds[1])
            self.set_link_tag(selection)
            edit_tag = user.get_id()[0] + ' ' + str(datetime.date.today())
            self.set_tag(edit_tag)
            self.add_text(name, bounds[0])
            self.set_link_tag(None)

    def tag_cb(self, texttag, widget, event, iter):
        """Process an event associated with a text tag."""
        if event.type == gtk.gdk.BUTTON_RELEASE:
            path = texttag.link
            if os.path.exists(path):
                if event.state & gtk.gdk.CONTROL_MASK:
                    path = os.path.dirname(path)
                if path:
                    if os.name == "nt":
                        os.startfile(path)
                    else:
                        subprocess.Popen(['gnome-open', path])
            elif urlparse.urlparse(path).scheme in ['http', 'https']:
                if os.name == "nt":
                    os.startfile(path)
                else:
                    subprocess.Popen(['gnome-open', path])
            else:
                top_level = widget.get_toplevel()
                top_level.open_link(path)

    def drop_text(self, data, point):
        """Add the text to the notes."""
        buffer = self.view.get_buffer()
        location = gtk.gdk.window_at_pointer()
        position = self.view.get_iter_at_location(location[1], location[2])
        if not position:
            position = buffer.get_end_iter()
        text = unicode(data.data, 'utf-8')
        text = text.replace(u'\x00', '')
        buffer.insert(position, text)

    def query_tooltip_cb(self, view, x, y, keyboard_mode, tooltip):
        """Determine what text to show in the tool tip."""
        iter = view.get_iter_at_location(x,y)
        tags = iter.get_tags()
        if tags:
            tags.reverse()
            text = [tag.get_property('name') for tag in tags if tag.get_property('name')]
            text = '\n'.join([string for string in text
                              if string.startswith('edit:') or string.startswith('link:')])
            text = text.replace('edit:', '')
            text = text.replace('link:', '')
            label = gtk.Label(text)
            tooltip.set_custom(label)
            return True
        return False

class Message(gtk.MessageDialog):
    """Simple wrapper around gtk message dialog widget."""

    @classmethod
    def show(cls, title, message, dialog_type):
        """Display a Message dialog dependent on the verbosity setting."""
        if (dialog_type == 'DEBUG' and config.get('verbosity') & 8) \
          or (dialog_type == gtk.MESSAGE_INFO and config.get('verbosity') & 4) \
          or (dialog_type == gtk.MESSAGE_WARNING and config.get('verbosity') & 2) \
          or (dialog_type == gtk.MESSAGE_ERROR and config.get('verbosity') & 1):
            dialog_type_map = {'DEBUG': gtk.MESSAGE_INFO}
            dialog_type =  dialog_type_map[dialog_type] if dialog_type in dialog_type_map else dialog_type
            cls(title, message, dialog_type)

    def __init__(self, title, message, dialog_type):
        """Creates and displays the dialog."""
        gtk.MessageDialog.__init__(self, None, 0, dialog_type, gtk.BUTTONS_CLOSE, message)
        self.dialog_type = dialog_type # Although this is unused it appears to be required to insure correct initialisation
        self.set_title(title)
        self.connect('response', self.response_cb)
        self.show_all()

    def response_cb(self, dialog, response_id):
        """Processes a response signal, the only response should be a click of the close button."""
        # Only expected response is a close.
        self.destroy()

class Confirmation(gtk.MessageDialog):
    """Simple wrapper around gtk message dialog widget for action confirmation."""

    def __init__(self, title, message, function, data=None):
        """Creates and displays the dialog."""
        self.function = function
        self.data = data
        gtk.MessageDialog.__init__(self, None, 0, gtk.MESSAGE_QUESTION, gtk.BUTTONS_YES_NO, message)
        self.set_title(title)
        self.connect('response', self.response_cb)
        self.show()

    def response_cb(self, dialog, response_id):
        """Processes a response signal, the only response should be a click of the close button."""
        if response_id == gtk.RESPONSE_YES:
            self.function(self.data)
        self.destroy()

class Event(object):
    """Class to manage event associated with a task transition"""
    #__slots__ = ['__id', 'time', 'prior_activity', 'subsequent_activity']
    ids = set()
    reminder = datetime.timedelta(minutes = 15.0)
    remind = True
    action = None

    def __init__(self, session, prior_activity=None, subsequent_activity=None, label=None):
        """Creates the event using the current time."""
        self.__id = None
        self.session = session
        if prior_activity:
            self.prior_activity = set([prior_activity])
        else:
            self.prior_activity = set()

        if subsequent_activity:
            self.subsequent_activity = set([subsequent_activity])
        else:
            self.subsequent_activity = set()
        self.__time = datetime.datetime.now()
        self.label = label

    def try_delete(self):
        """Delete this event if it is no longer used."""
        if not self.label and not self.prior_activity and not self.subsequent_activity:
            Message.show("Delete Event", "Deleting event %s." % (self), 'DEBUG')
            self.session.events.remove(self)

    def get_id(self):
        """Returns a unique id for the event, generating an id if necessary."""
        if not self.__id:
            self.set_id(1)
        return self.__id

    def set_id(self, id):
        """Sets the id for the event."""
        if id in self.session.event_ids:
            self.set_id(id + 1)
        else:
            self.__id = id
            self.session.event_ids.add(id)

    def get_time(self):
        """Returns time for this event."""
        return self.__time

    def set_time(self, time):
        """Changes the time for this event and updates the associated activities."""
        self.__time = time
        if time < datetime.datetime.now():
            self.remind = False
        else:
            self.remind = True
        [prior.update_duration() for prior in self.prior_activity]
        [subsequent.update_duration() for subsequent in self.subsequent_activity]

    time = property(get_time, set_time)

    def __str__(self):
        """Return a string representation of the event."""
        result = ''
        if self.label:
            result = self.label
        else:
            result = self.time.strftime(config.get('time_format'))
        return result

    def info(self):
        """Return a string containing the details of the event."""
        result = "Event %s: (" % (self.get_id())
        for activity in self.prior_activity:
            try:
                result += "%d, " % (activity.get_id())
            except:
                result += "%d, " % (activity)
        if self.prior_activity:
            result = result[:-2]
        result += ') ('
        for activity in self.subsequent_activity:
            try:
                result += "%d, " % (activity.get_id())
            except:
                result += "%d, " % (activity)
        if self.subsequent_activity:
            result = result[:-2]
        result += ') ' + self.time.isoformat()
        return result

    def to_xml(self, document):
        """Create an XML node for this event."""
        event_node = document.createElement('event')
        event_node.setAttribute('id', str(self.get_id()))

        time_node = document.createElement('time')
        time_node.appendChild(document.createTextNode(self.time.isoformat()))
        event_node.appendChild(time_node)

        if self.label:
            subsequent_activity_node = document.createElement('label')
            subsequent_activity_node.appendChild(document.createTextNode(self.label))
            event_node.appendChild(subsequent_activity_node)

        if self.prior_activity:
            prior_activity_node = document.createElement('prior')
            the_ids = [str(activity.get_id()) for activity in self.prior_activity]
            prior_activity_node.appendChild(document.createTextNode(' '.join(the_ids)))
            event_node.appendChild(prior_activity_node)

        if self.subsequent_activity:
            subsequent_activity_node = document.createElement('subsequent')
            the_ids = [str(activity.get_id()) for activity in self.subsequent_activity]
            subsequent_activity_node.appendChild(document.createTextNode(' '.join(the_ids)))
            event_node.appendChild(subsequent_activity_node)

        return event_node

    def from_xml(self, node):
        """Load this event from an XML node."""
        self.set_id(int(node.getAttribute('id')))
        for child_node in node.childNodes:
            if child_node.tagName == 'time':
                self.time = string_to_time(child_node.firstChild.nodeValue)
            elif child_node.tagName == 'label':
                self.label = child_node.firstChild.nodeValue
            elif child_node.tagName == 'prior':
                self.prior_activity = set([int(id) for id in child_node.firstChild.nodeValue.split()])
            elif child_node.tagName == 'subsequent':
                self.subsequent_activity = set([int(id) for id in child_node.firstChild.nodeValue.split()])

class Now:
    """Special object representing the current time."""

    def __init__(self, session, label = None):
        """Initialises the 'Now' event without adding it to the events list."""
        self.__id = None
        self.session = session
        self.label = label
        self.subsequent_activity = set()

    def __str__(self):
        """Return a string representation of the event."""
        return 'Now'

    def try_delete(self):
        """Interface required to match 'event' objects."""
        pass

    def get_id(self):
        """Returns id of 0 to indicate 'Now' event"""
        return 0

    def get_time(self):
        """Returns current time."""
        return datetime.datetime.now()

    time = property(get_time)

    def get_prior(self):
        """Return the current activity."""
        if self.session.current_activity:
            return set([self.session.current_activity])
        return set()

    prior_activity = property(get_prior)

class Activity(object):
    """Class to manage an activity which combines a task with start and end times"""
    ids = set()
    __allocation = None
    action = None

    def __init__(self, session, task, start=False):
        """Create an activity associated with the given task."""
        self.__id = 0
        self.session = session
        self.task = task
        self.outlook_id = None
        if start:
            self.start_event = self.session.new_event()
            self.start_event.subsequent_activity.add(self)
            if task and not task.start_date:
                task.start_date = self.start_event.time
            self.end_event = Now(session)
        else:
            self.start_event = None
            self.end_event = None
        self.duration = datetime.timedelta()

    def delete(self):
        """Removes this activity and the references to it in the event."""
        if self.task:
            self.task.update_hours(-self.duration)
        if self.start_event:
            if self in self.start_event.subsequent_activity:
                self.start_event.subsequent_activity.remove(self)
                self.start_event.try_delete()
                self.start_event = None
            else:
                Message.show('Inconsistent State',
                             'Inconsistent start event for activity %d (%s)' % (self.get_id(), self.task.name),
                             gtk.MESSAGE_WARNING)
        if self.end_event:
            if self in self.end_event.prior_activity:
                self.end_event.prior_activity.remove(self)
                self.end_event.try_delete()
                self.end_event = None
            else:
                Message.show('Inconsistent State',
                             'Inconsistent start event for activity %d (%s)' % (self.get_id(), self.task.name),
                             gtk.MESSAGE_WARNING)
        self.session.activities.remove(self)

    def get_id(self):
        """Returns a unique id for the activity, generating an id if necessary."""
        if not self.__id:
            self.set_id(1)
        return self.__id

    def set_id(self, id):
        """Sets the id for the activity."""
        if id in self.session.activity_ids:
            self.set_id(id + 1)
        else:
            self.__id = id
            self.session.activity_ids.add(id)

    def set_task(self, task):
        """Update the task associated with this activity, adjusting the hours
           allocated to the task.

        """
        if task == self.task:
            return

        if self.task:
            self.task.update_hours(-self.get_duration())
        if task:
            task.update_hours(self.get_duration())
            if not task.start_date:
                task.start_date = self.start_event.time
        self.task = task

    def get_allocation(self):
        """Returns the allocation for this activity or the associated task."""
        if self.__allocation != None:
            return self.__allocation
        elif self.task:
            return self.task.allocation
        else:
            return 1.0

    def set_allocation(self, allocation):
        """Sets the allocation for this activity."""
        self.__allocation = allocation

    allocation = property(get_allocation, set_allocation)

    def change_start(self, new_start):
        """Updates the start event."""
        old_start = self.start_event
        if old_start == new_start:
            return
        self.start_event = new_start
        new_start.subsequent_activity.add(self)
        self.update_duration()
        if old_start:
            old_start.subsequent_activity.remove(self)
            old_start.try_delete()
        if self.task and not self.task.start_date:
            self.task.start_date = new_start.time

    def change_end(self, new_end):
        """Updates the end event."""
        old_end = self.end_event
        if old_end == new_end:
            return
        self.end_event = new_end
        if new_end:
            new_end.prior_activity.add(self)
            self.update_duration()
        if old_end and old_end.get_id():
            old_end.prior_activity.remove(self)
            old_end.try_delete()

    def swap(self):
        """Swaps the start and end events for this activity."""
        self.start_event, self.end_event = self.end_event, self.start_event
        self.start_event.subsequent_activity.add(self)
        self.start_event.prior_activity.remove(self)
        self.end_event.prior_activity.add(self)
        self.end_event.subsequent_activity.remove(self)

    def split(self, event):
        """Splits this activity at the supplied event"""
        new_activity = None
        end = self.end_event.time if self.end_event else datetime.datetime.now()
        if event.time >= self.start_event.time and event.time <= end:
            old_start = self.start_event
            new_activity = self.session.new_activity(self.task, old_start, event)
            self.change_start(event)
        else:
            Message.show("Invalid Time",
                         "Attempt to split activity %s by event outside duration." % (self.task.name),
                         gtk.MESSAGE_WARNING)
        return new_activity

    def occupy_full_day(self):
        """Adjusts activity to last all day.""" 
        start = datetime.time(hour = int(config.get('day_start')[:2]),
                              minute = int(config.get('day_start')[3:]),
                              second = 0,
                              microsecond = 0)
        start = datetime.datetime.combine(self.start_event.time.date(), start)
        duration = string_to_timedelta(config['hours_per_week'] / config['days_per_week'])
        end = start + duration
        if not self.start_event.prior_activity:
            self.start_event.time = start
        else:
            new_event = self.session.new_event()
            new_event.time = start
            self.change_start(new_event)
            
        if not self.end_event.subsequent_activity:
            self.end_event.time = end
        else:
            new_event = self.session.new_event()
            new_event.time = end
            self.change_end(new_event)

    def get_duration(self):
        """Returns the duration of this activity."""
        duration = datetime.timedelta()
        if self.end_event:
            end_time = self.end_event.time
        elif self.is_active():
            end_time = datetime.datetime.now()
        elif self.start_event:
            end_time = self.start_event.time
        if self.start_event:
            duration = end_time - self.start_event.time
        return duration

    def update_duration(self):
        """Updates the duration of this activity changing task allocated hours."""
        initial_duration = self.duration
        self.duration = self.get_duration()
        if self.task:
            self.task.update_hours(self.duration - initial_duration)

    def is_active(self):
        """Returns true if this activity is the current activity."""
        return self == self.session.current_activity

    def __str__(self):
        """Display the data for this activity."""
        start = str(self.start_event) if self.start_event else None
        end = str(self.end_event) if self.end_event else None
        duration = self.get_duration()
        if duration:
            hours = int(duration.seconds / 3600)
            minutes = int((duration.seconds - hours * 3600) / 60)
            seconds = int(duration.seconds - hours * 3600 - minutes * 60)
            duration = "%s hr %s min %s sec" % (hours, minutes, seconds)
        else:
            duration = None
        return "Activity %d: %s, %s to %s (%s)" % (self.__id, self.task.name, start, end, duration)

    def to_xml(self, document):
        """Create an XML node for this event."""
        activity_node = document.createElement('activity')
        activity_node.setAttribute('id', str(self.get_id()))

        if self.task:
            task_node = document.createElement('task')
            task_node.appendChild(document.createTextNode(str(self.task.get_id())))
            task_node.appendChild(document.createComment('"%s"' % (self.task.full_name())))
            activity_node.appendChild(task_node)

        if self.start_event:
            start_node = document.createElement('start')
            start_node.appendChild(document.createTextNode(str(self.start_event.get_id())))
            activity_node.appendChild(start_node)

        if self.end_event and self.end_event.get_id() != 0:
            end_node = document.createElement('end')
            end_node.appendChild(document.createTextNode(str(self.end_event.get_id())))
            activity_node.appendChild(end_node)

        if self.outlook_id:
            id_node = document.createElement('outlook_id')
            id_node.appendChild(document.createTextNode(str(self.outlook_id)))
            activity_node.appendChild(id_node)

        if self.allocation < 1.0:
            allocation_node = document.createElement('allocation')
            allocation_node.appendChild(document.createTextNode(str(self.allocation)))
            activity_node.appendChild(allocation_node)

        return activity_node

    def from_xml(self, node):
        """Load this activity from an XML node."""
        self.set_id(int(node.getAttribute('id')))
        for child_node in node.childNodes:
            if child_node.tagName == 'task':
                self.task = int(child_node.firstChild.nodeValue)
            elif child_node.tagName == 'start':
                self.start_event = int(child_node.firstChild.nodeValue)
            elif child_node.tagName == 'end':
                self.end_event = int(child_node.firstChild.nodeValue)
            elif child_node.tagName == 'outlook_id':
                self.outlook_id = child_node.firstChild.nodeValue
            elif child_node.tagName == 'allocation':
                self.allocation = float(child_node.firstChild.nodeValue)

class Time_Metrics(object):
    """Class to manage metrics of a task measured by effort."""
    def __init__(self, name, required):
        """Initialise the data required for measurement."""
        self.current = 0
        self.required = required
    
class Quanitity_Metrics(object):
    """Class containing metrics associated with quantity measurements."""
    
    def __init__(self, name, required):
        """Initialise the data required for measurement."""
        self.name = name
        self.current = 0
        self.required = required

    def get_progress(self):
        """Returns the progress determined by this metric."""
        return 100.0 * float(self.current) / float(self.required) 

class Task(object):
    """Class to manage data associated with a single task."""
    #__slots__ = ['__id', 'name', 'booking_number', 'priority', 'link', 'state', 'parent', 'children']

    states = ['Deleted',    # Old task maintained to avoid breaking references
              'New',        # Initial state of newly created task
              'Pending',    # Not yet started
              'Active',     # Started but not completed
              'Scheduled',  # Work to be carried out at predetermined time e.g. a meeting
              'Parent',     # Task for which all work is carried out in child tasks 
              'Inactive',   # Dormant task that may be resurrect 
              'Complete']   # Completed task
              
    measurement_type = ['Time',     # Completion based purely by amount of time spent
                        'Quantity'] # Completion determined be quantity of output

    xml_data = {'name':('name', None, None, True),
                'booking_number':('booking_number', None, None, True),
                'start':('_start_date', datetime.datetime.isoformat, string_to_date, True),
                'end':('_end_date', datetime.datetime.isoformat, string_to_date, True),
                'allocated':('allocated_time', get_decimal_hours, string_to_timedelta, True),
                'estimated':('estimated_time', get_decimal_hours, string_to_timedelta, True),
                'priority':('priority', None, int, True),
                'notes':('notes', None, None, True),
                'state':('state', None, None, True),
                'outlook_id':('outlook_id', None, None, True)}
    ids = set()
    _start_date = None
    _end_date = None

    def __init__(self, parent = None):
        """Create the task"""
        self.__id = None
        self.state = 'New'
        self.name = 'New Task'
        self.booking_number = None
        self.metric = 'Quantity'
        self.allocated_time = datetime.timedelta()
        self.estimated_time = datetime.timedelta()
        self.output = 0
        self.required_output = 1
        self.priority = 0
        self.allocation = 1.0
        self.outlook_id = ''
        self.notes = ''
        self.parent = parent
        self.recurring_activities = []
        self.children = []

    def get_start_date(self):
        """Returns the start date of this task or the earliest start date of a child if not set."""
        start_date = self._start_date
        if not start_date and self.children:
            for child in self.children:
                child_start = child.start_date
                if (not start_date) or (child_start and child_start < start_date):
                   start_date = child_start
        return start_date

    def set_start_date(self, date):
        """Sets the start date of this task."""
        self._start_date = date

    start_date = property(get_start_date, set_start_date)

    def get_end_date(self):
        """Returns the end date of this task or the latest end date of a child if not set."""
        end_date = self._end_date
        if not end_date and self.children:
            for child in self.children:
                child_end = child.end_date
                if (not end_date) or (child_end and child_end > end_date):
                   end_date = child_end
        return end_date

    def set_end_date(self, date):
        """Sets the end date of this task."""
        self._end_date = date

    end_date = property(get_end_date, set_end_date)

    def get_state(self):
        """Returns the state as a string."""
        return Task.states[self._state]

    def set_state(self, new_state):
        """Sets the state of the task."""
        try:
            self._state = Task.states.index(new_state.capitalize())
        except:
            self._state = new_state

        if self._state == Task.states.index('Active') and not self.start_date:
            self.start_date = datetime.datetime.now()
        if self._state == Task.states.index('Complete') and not self.end_date:
            self.end_date = datetime.datetime.now()
            if self.parent and self in self.parent.children and self.parent.state == 'Parent':
                self.parent.get_progress()

    state = property(get_state, set_state)

    def get_progress(self):
        """Determine if this task is complete based on measurement and children state."""
        # TODO: Implement measurement functionality
        progress = 0.0
        if config.get('inactive_is_complete'):
            completion_states = ['Complete', 'Deleted', 'Inactive']
        else:
            completion_states = ['Complete', 'Deleted']

        if self.state == 'Parent':
            for child in self.children:
                progress += child.get_progress() if child.state not in completion_states else 100.0
            if self.children:
                progress /= len(self.children)
            else:
                Message.show('Task State Error',
                             'Parent task %s has no children' % (self.name),
                              gtk.MESSAGE_WARNING)
        else:
            if self.state in completion_states:
                progress = 100.0
            elif self.metric == 'Time':
                if self.allocated >= self.estimated:
                    progress = 100.0
                else:
                    allocated = get_decimal_hours(self.allocated_time)
                    estimated = get_decimal_hours(self.estimated_time)
                    progress = 100.0 * allocated / estimated
            else:
                if self.output == self.required_output:
                    progress = 100.0
                else:
                    progress = 100.0 * (float(self.output) /
                                        float(self.required_output))
             
        return progress

    def is_runable(self):
        """Determine if this task is active."""
        if self.state in ['Pending', 'Active', 'Scheduled']:
            return True
        else:
            return False

    def has_runable_child(self):
        """Returns true if this task or any of it's children are runable."""
        if self.is_runable():
            return True
        for child in self.children:
            if child.has_runable_child():
                return True
        return False

    def full_name(self):
        """Returns the full name of the task, including parents."""
        if self.parent:
            full_name = "%s/%s" %(self.parent.full_name(), self.name)
        else:
            full_name = self.name
        return full_name

    def _get_booking_number(self):
        """Return the booking number prepeded with parent booking number."""
        booking_number = self.booking_number
        if not booking_number:
            booking_number = self.name
        if self.parent:
            parent_booking_number = self.parent._get_booking_number()
            if parent_booking_number:
                booking_number = "%s/%s" % (parent_booking_number, booking_number)
        return booking_number

    def get_booking_number(self):
        """Return the booking number prepeded with parent booking number."""
        booking_number = self._get_booking_number()
        booking_number = booking_number.replace('/-', '')
        if booking_number.startswith('/'):
            booking_number = booking_number[1:]

        # Catch pathological cases of all '/-' or single '/'
        if not booking_number:
            booking_number = activity.task.full_name()

        return booking_number

    def is_billable(self):
        """Return a boolean indicating whether the task is billable."""
        booking_number = self.get_booking_number()
        return not booking_number.startswith('#')

    def get_hours_subtotal(self):
        """Returns the sum of allocated hours for this task and it's children."""
        total = get_decimal_hours(self.allocated_time)
        for child in self.children:
            total += child.get_hours_subtotal()
        return total

    def show(self, level = 0):
        """Display the task and sub tasks."""
        print ("%s%s" % (' ' * level, self.full_name()))
        for child in self.children:
            child.show(level + 1)

    def get_id(self):
        """Returns a unique id for the task, generating an id if necessary."""
        if not self.__id:
            self.set_id(1)
        return self.__id

    def set_id(self, id):
        """Sets the id for the task."""
        if id in Task.ids:
            self.set_id(id + 1)
        else:
            self.__id = id
            Task.ids.add(id)

    def update_hours(self, delta):
        """Update the hours allocated to this task."""
        self.allocated_time += delta
        if self.allocated_time < datetime.timedelta():
            self.allocated_time = datetime.timedelta()
        if self.parent:
            self.parent.update_hours(delta)

    def to_xml(self, document):
        """Create an XML node for this task."""
        task_node = document.createElement('task')
        task_node.setAttribute('id', str(self.get_id()))

        for node_name in self.xml_data.keys():
            attribute, to_xml, from_xml, dummy = self.xml_data[node_name]
            value = getattr(self, attribute)
            if value:
                node = document.createElement(node_name)
                if to_xml:
                    child_node = document.createTextNode(str(to_xml(value)))
                else:
                    child_node = document.createTextNode(str(value))
                node.appendChild(child_node)
                task_node.appendChild(node)
        for outlook_id in self.recurring_activities:
            id_node = document.createElement('recurring')
            id_node.appendChild(document.createTextNode(str(outlook_id)))
            task_node.appendChild(id_node)

        for child in self.children:
            task_node.appendChild(child.to_xml(document))

        return task_node

    def from_xml(self, node):
        """Load this task from an XML node."""
        self.set_id(int(node.getAttribute('id')))
        for child_node in node.childNodes:
            if child_node.tagName in self.xml_data:
                attribute, to_xml, from_xml, dummy = self.xml_data[child_node.tagName]
                value = child_node.firstChild.nodeValue
                if from_xml:
                    value = from_xml(value)
                setattr(self, attribute, value)
            elif child_node.tagName == 'task':
                new_task = Task(self)
                if new_task:
                    new_task.from_xml(child_node)
                    self.children.append(new_task)
            elif child_node.tagName == 'recurring':
                self.recurring_activities.append(child_node.firstChild.nodeValue)

class Task_Manager:
    """Management of tasks."""
    tasks = []
    notes = ''

    def __init__(self, update_gui):
        """Initialise the object loading the tasks list."""
        self.update_gui = update_gui
        self.recent_tasks = []
        if config.get('task_file'):
            self.from_xml(config.get('data_dir') + config.get('task_file'))

    def find_by_path(self, path):
        """Returns the task based on the full name."""
        sep = '/'
        if path.count('-') > path.count('/'):
            sep = '-'
        path = path.strip(sep)
        first = path.find(sep)
        sub_path = None
        if first > 0:
            task = find(lambda task: task.name == path[:first], self.tasks)
            sub_path = path[first:]
        else:
            task = find(lambda task: task.name == path, self.tasks)

        if task and sub_path:
            task = self.find_descendent(task, sub_path)
        return task

    def find_task(self, subject, creation_state=None):
        """Try to find the task associated with the given subject."""
        results = self.find_by_subject(subject)
        high = max(results.values())
        matches = [task for task in results.keys() if results[task] == high]

        if len(matches) == 1:
            result = matches[0]
        elif creation_state:
            parents = set([task for task in results.keys() if results[task] == max])
            decendents = set([task for task in parents if task.parent in parents])
            parents = parents - decendents

            if len(parents) == 1 and max > 0.4:
                parent = parents.pop()
                name = []
                words = [word for word in subject.split()]
                lower_words = [word.lower() for word in subject.split()]
                for word in short_name[parent]:
                    if word in words:
                        name.append(word)
                    else:
                        name.append(words[lower_words.index(word)])
                name = ' '.join(name)
                new_task = self.new_task(name, parent=parent)
            else:
                new_task = self.new_task(subject)
            new_task.state = creation_state
            result = new_task
        else:
            result = None
        return result

    def determine_parent(self, subject):
        """Return a guess of the parent task based on the given subject."""
        results = self.find_by_subject(subject)
        parent = max(results.items(), key=operator.itemgetter(1))[0]
        if results[parent] > 0:
            while parent.parent and results[parent] - results[parent.parent] < 0.4:
                parent = parent.parent
            parent_name = parent.full_name().lower().replace('/', ' ')
            parent_name = parent_name.split()

            subject_string = subject.replace('/', ' ')
            subject_string = subject_string.replace(' - ', ' ')
            subject_words = [word for word in subject_string.split()]

            for word in [w for w in subject_words]:
                if word.lower() in parent_name:
                    subject_words.remove(word)
    
            task_name = ' '.join(subject_words)
        else:
            parent = None
            task_name = subject
        return (parent, task_name)

    def find_by_subject(self, subject):
        """Returns a list of tasks associated with the given subject."""
        words = {}

        # Compile dictionary of words and the number of times each is used
        for task in expand_list(self.tasks):
            name = task.name.replace('/', ' ')
            for word in name.split():
                if word == '-':
                    continue
                lower = word.lower()
                if lower in words:
                     words[lower] += 1
                else:
                     words[lower] = 1

        level = 0.0
        results = {}
        for task in expand_list(self.tasks):
            results[task] = 0
            if subject.lower().endswith(task.name.lower()):
                results[task] = 1
            pos = subject.find(task.name)
            if pos != -1:
                #if subject.startswith(task.name) == 0:
                #    results[task] += 0.5
                #elif subject.find(task.name) > 0:
                #    results[task] += 0.4
                #while 0:
                parent = task
                while parent:
                    if subject.find(parent.name) == 0:
                        results[task] += 0.5
                    elif subject.find(parent.name) > 0:
                        results[task] += 0.4
                    parent = parent.parent
            else:                    
                subject_string = subject.replace('/', ' ')
                subject_words = [word.lower() for word in subject_string.split()]
                results[task] = level
                if task.parent:
                    results[task] += results[task.parent]
                name = task.name.replace('/', ' ')
                for word in name.split():
                    if word == '-':
                        continue
                    lower = word.lower()
                    if lower in subject_words:
                        results[task] += 1.0 / (words[lower] * (subject_words.index(lower) + 1.0))
        return results

    def guess_by_subject(self, subject):
        """Returns a list of tasks associated with the given subject."""
        words = {}

        # Compile dictionary of words and the number of times each is used
        for task in expand_list(self.tasks):
            for word in task.name.split():
                lower = word.lower()
                if lower in words:
                     words[lower] += 1
                else:
                     words[lower] = 1

        subject_words = [word.lower() for word in subject.split()]
        level = 0.0
        results = {}
        for task in expand_list(self.tasks):
            results[task] = level
            for word in task.name.split():
                lower = word.lower()
                if lower in subject_words:
                    results[task] += 1.0 / (words[lower] * (subject_words.index(lower) + 1.0))

        return results

    def find_by_outlook_id(self, outlook_id):
        """Returns the task with the given outlook id, or None if no match."""
        task = find(lambda task : task.outlook_id == outlook_id, expand_list(self.tasks))
        return task

    def task_activated(self, task):
        """Append the task to the list of recent tasks."""
        while task in self.recent_tasks:
           self.recent_tasks.remove(task)
        self.recent_tasks.insert(0, task)
        self.recent_tasks = self.recent_tasks[:config.get('len_recent_tasks')]

    def delete_task(self, task):
        """Deletes the task allocating any events to the parent."""
        task.state = 'deleted'
        if task.parent:
            task.parent.children.remove(task)
        else:
            self.tasks.remove(task)
        self.to_xml(config.get('data_dir') + config.get('task_file'))
        window.update_task_menu()

    def new_child_task_cb(self, item, task):
        """Displays a new task dialog with the specified task as parent."""
        Task_Dialog(self, None, parent=task)

    def make_heirarchy(self, task, function, sort=False, all_tasks=False):
        """Recursively create a heirarchy based on the task structure."""
        entry = None
        sub_menu = []
        children = task.children
        if sort:
            children.sort()
        if config.get('submenus_on_top'):
            sub_menus = [child for child in children if child.children]
            elements = [child for child in children if not child.children]
            children = sub_menus + elements
        for child in children:
            menu_entry = self.make_heirarchy(child, function, sort, all_tasks)
            if menu_entry:
                sub_menu.append(menu_entry)
        if sub_menu:
            if task.is_runable() or all_tasks:
                sub_menu = [("%s" % (task.name), function, task),
                           ('---', '---')] + sub_menu
            sub_menu += [('---','---'),
                         ('New Sub Task', self.new_child_task_cb, task)]
            entry = (">>>%s" % (task.name), sub_menu)
        elif task.is_runable() or all_tasks:
            entry = ("%s" % (task.name), function, task)
        return entry

    def make_flat(self, task, function, sort=False):
        """Recursively create a heirarchy based on the task structure."""
        entry = []
        if task.is_runable():
            entry = [("%s" % (task.full_name()), function, task)]
        if task.children:
            children = task.children
            if sort:
                children.sort()
            for child in children:
                entry.extend(self.make_flat(child, sort))
        return entry

    def create_menu_definition(self, function, use_sub_menus=True, sort=False, all_tasks=False):
        """Creates a menu definition for use be the build_menu function."""
        menu = []
        if use_sub_menus:
            # Convert the task names into a tree structure
            task_list = self.tasks
            if sort:
                task_list.sort()
            if config.get('submenus_on_top'):
                sub_menus = [child for child in task_list if child.children]
                elements = [child for child in task_list if not child.children]
                task_list = sub_menus + elements
            
            for task in task_list:
                menu_entry = self.make_heirarchy(task, function, sort, all_tasks)
                if menu_entry:
                    menu.append(menu_entry)
        else:
            # Create flat structure
            tasks = self.tasks
            if sort:
                tasks.sort()
            for task in tasks:
                menu.extend(self.make_flat(task, function, sort))
        return menu

    def new_task(self, path, prior=None, parent=None):
        """Create a new task."""
        if prior:
            if prior.parent:
                task_list = prior.parent.children
            else:
                task_list = self.tasks
            task = Task(prior.parent)
            task.name = path
            task_list.insert(task_list.index(prior), task)
        elif parent:
            task = Task(parent)
            task.name = path
            parent.children.append(task)
        elif path:
            tasks = self.tasks
            task = None
            for element in path.split('/'):
                child = find(lambda task: task.name == element, tasks)
                if not child:
                    child = Task(task)
                    child.name = element
                    tasks.append(child)
                task = child
                tasks = task.children
        else:
            task = Task()
            task.name = path
            self.tasks.append(task)

        if task.state != 'New':
            Message.show("Duplicate name!", "Attempted to add task %s twice." % (path), gtk.MESSAGE_WARNING)
            return None
        else:
            task.state = 'Pending'
            self.write_task_list()
        return task

    def reparent(self, task, new_parent):
        """Moves a task within the tree."""
        if task.parent:
            current_location = task.parent.children
        else:
            current_location = self.tasks

        if new_parent:
            new_location = new_parent.children
        else:
            new_location = self.tasks

        if new_parent == task.parent and task in new_location:
            return

        if task in current_location:
            current_location.remove(task)
        task.parent = new_parent
        new_location.append(task)
        self.write_task_list()

    def update_children(self, parent, new_list):
        """Updates the list of child tasks for the given parent."""
        if parent:
            parent.children = new_list
        else:
            self.tasks = new_list
        self.write_task_list()

    def find_child_by_name(self, name):
        """Returns the child task by name."""
        return find(lambda task: task.name == name, self.children)

    def find_descendent(self, task, path):
        """Returns the descendent task from the path."""
        path = path.strip('/')
        first = path.find('/')
        sub_path = None
        if first > 0:
            descendent = find(lambda task: task.name == path[:first], task.children)
            sub_path = path[first:]
        else:
            descendent = find(lambda task: task.name == path, task.children)

        if descendent and sub_path:
            descendent = self.find_descendent(descendent, sub_path)

        return descendent

    def show(self):
        """Display the times for current tasks."""
        for task in self.tasks:
            task.show()

    def write_task_list(self):
        """Simple wrapper round the xml output of the task list."""
        filename = config.get('data_dir') + config.get('task_file')
        self.to_xml(filename)
        if config.get('daily_backup'):
            directory = os.path.expanduser(config.get('data_dir') + '/backup')
            if not os.path.exists(directory):
                os.makedirs(directory)
            filename = os.path.expanduser(filename)
            backup = directory + '/' + config.get('task_file')
            path, extension = os.path.splitext(backup)
            day = datetime.date.today().strftime('%A')
            backup = "%s_%s%s" % (path, day, extension)
            shutil.copy(filename, backup)
        self.update_gui()

    def to_xml(self, filename):
        """Writes the task list data to file."""
        filename = os.path.expanduser(filename)

        document = xml.dom.minidom.getDOMImplementation().createDocument(None, None, None)

        root_node = document.createElement('tasks')

        for task in self.tasks:
            root_node.appendChild(task.to_xml(document))

        if self.recent_tasks:
            recent = ' '.join([str(task.get_id()) for task in self.recent_tasks if task])
            recent_node = document.createElement('recent_tasks')
            recent_node.appendChild(document.createTextNode(recent))
            root_node.appendChild(recent_node)
        document.appendChild(root_node)

        if self.notes:
            note_node = document.createElement('notes')
            note_node.appendChild(document.createTextNode(self.notes))
            root_node.appendChild(note_node)

        if not os.path.exists(os.path.dirname(filename)):
            os.makedirs(os.path.dirname(filename))
        file = open(filename, 'w')
        file.write(document.toprettyxml())
        file.close()

    def outlook_import(self, namespace):
        """Imports tasks from outlook."""
        results = ''
        task_folders = []
        for folder in namespace.Folders:
            try:
                task_folders.append(folder.Folders['Tasks'])
                results += 'Adding tasks from %s folder\n' % (folder.name)
            except:
                results += 'Skipping folder: %s\n' % (folder.Name)

        for folder in task_folders:
            for item in folder.items:
                outlook_id = item.EntryID
                task = self.find_by_outlook_id(outlook_id)
                if not task:
                    task = self.find_task(item.Subject, 'Pending')
                    task.outlook_id = item.EntryID
                start_date = str(item.StartDate)
                if start_date == '01/01/01 00:00:00':
                    start_date = str(item.CreationTime)
                task.start_date = string_to_date(start_date, True)
                body = item.Body
                if body not in task.notes:
                    task.notes += body
                billing_info = item.BillingInformation
                if not task.booking_number:
                    task.booking_number = billing_info
                else:
                    task.notes += '\nBilling Information: %s\n' % (billing_info)
                if item.Complete:
                    task.state = 'complete'
                    completed = str(item.DateCompleted)
                    task.end_date = string_to_date(completed, True)

        return results

    def duration_from_xml(self, node):
        """Load the duration data from an XML node."""
        start_event = None
        end_event = None
        for child_node in node.childNodes:
            if child_node.tagName == 'task':
                task = int(child_node.firstChild.nodeValue)
            elif child_node.tagName == 'start_event':
                start_event = int(child_node.firstChild.nodeValue)
            elif child_node.tagName == 'end_event':
                end_event = int(child_node.firstChild.nodeValue)
        return (task, start_event, end_event)

    def from_xml(self, filename):
        """Load an xml task file."""
        result = False
        pattern = re.compile("\s*(<|>)\s*")
        filename = os.path.expanduser(filename)
        activities = []

        if os.path.exists(filename):
            file = open(filename)
            xml_string = file.read()
            file.close()

            # Remove pretty printing
            xml_string = pattern.sub(r'\1', xml_string)
            document = xml.dom.minidom.parseString(xml_string)
            for node in document.documentElement.childNodes:
                if node.tagName == 'task':
                    new_task = Task()
                    new_task.from_xml(node)
                    self.tasks.append(new_task)
                if node.tagName == 'recent_tasks':
                    if node.firstChild:
                        recent = node.firstChild.nodeValue.split()
                        self.recent_tasks = [find(lambda task: task.get_id() == int(task_id), self.tasks) for task_id in recent]
                if node.tagName == 'notes':
                    if node.firstChild:
                        self.notes = node.firstChild.nodeValue
        return True

class Activity_Manager(object):
    """Management of activities for a week."""
    sessions = {}
    _current = None

    def __init__(self, task_manager, update_gui, date=None, filename=None):
        """Initialise the object using the current date."""
        self.current_activity = None
        self.previous_task = None
        self.activities = []
        self.activity_ids = set()
        self.events = []
        self.event_ids = set()
        self.task_manager = task_manager
        self.update_gui = update_gui
        if not date:
            date = datetime.datetime.now()
        self.week_start = date.replace(hour=0, minute=0, second=0, microsecond=0)
        self.week_start -= datetime.timedelta(date.weekday())
        self.week_end = self.week_start + datetime.timedelta(7) - datetime.timedelta(microseconds=1)

        if not filename:
            if config.get('session_file'):
                path, extension = os.path.splitext(config.get('data_dir') + config.get('session_file'))
                filename = "%s_%s%s" % (path, self.week_end.date().strftime('%Y-%m-%d'), extension)
        self.from_xml(filename)                
        self.sessions[self.week_end] = self

    def get_current(self):
        """Returns the current activity."""
        return self._current
        
    def set_current(self, activity):
        """Sets the current activity, updating the end event on the previous activity if necessary."""
        if self._current and self._current != activity:
            if activity:
                self._current.change_end(activity.start_event)
            else:
                self._current.change_end(self.new_event())        
        self._current = activity
        if activity:
            self.task_manager.task_activated(activity.task)
            activity.start_event.remind = False
        
    current_activity = property(get_current, set_current)

    def is_current(self):
        """Returns a boolean indicating whether it's valid to start a task in
           this session, i.e. is the current time between the session start and
           end times.

           """
        result = False
        now = datetime.datetime.now()
        if now > self.week_start and now < self.week_end:
            result = True
        return result

    def activate_cb(self, widget, task):
        """Starts the specified task."""
        self.start_task(task)

    def outlook_import_cb(self, widget, function=None):
        """Imports tasks from the outlook calendar."""
        if not os.name == 'nt':
            return

        results = ""
        # Attempt to connect to outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        if config.get('import_tasks'):
            results += self.task_manager.outlook_import(namespace)

        # Outlook seems to enjoy moving/renaming folders and not giving an easy way to find them!
        calendars = []
        for folder in namespace.Folders:
            try:
                calendars.append(folder.Folders['Calendar'])
            except:
                results += "Skipping folder: %s\n" % (folder.Name)

        week_start = self.week_start
        week_end = week_start + datetime.timedelta(days = 7)
        week_start = pywintypes.Time(week_start)
        week_end = pywintypes.Time(week_end)
        filter = u"[Start] >= '%s' And [End] <= '%s'" % (week_start.Format("%x %I:%M %p"), week_end.Format("%x %I:%M %p"))
        start_filter = u"[Start] >= '%s'" % (week_start.Format("%d/%m/%y %I:%M %p"))
        end_filter = u"[End] <= '%s'" % (week_end.Format("%d/%m/%y %I:%M %p"))
        for calendar in calendars:
            appointments = calendar.Items
            appointments.IncludeRecurrences = True
            appointments.sort("Start")
            for item in appointments.Restrict(start_filter).restrict(end_filter):
                start_time = datetime.datetime(item.start.year,
                                               item.start.month,
                                               item.start.day,
                                               item.start.hour,
                                               item.start.minute,
                                               item.start.second,
                                               item.start.msec)
                end_time = datetime.datetime(item.end.year,
                                             item.end.month,
                                             item.end.day,
                                             item.end.hour,
                                             item.end.minute,
                                             item.end.second,
                                             item.end.msec)
                subject = item.subject.replace('Updated: ', '')

                task = None
                # Recurring tasks are found by outlook Entry ID
                if item.IsRecurring:
                    task = find(lambda task : item.EntryID in task.recurring_activities,
                                expand_list(self.task_manager.tasks))

                # Non recurring tasks or new recurring tasks are found by subject
                if not task:
                    parent, task_name = self.task_manager.determine_parent(subject)
                    if task_name:
                        task = self.task_manager.new_task(task_name, parent=parent)
                        task.state = "Scheduled"
                        results += "Creating new task: %s" % (task.full_name())
                    else:
                        task = parent
                        task.state = "Scheduled"
                        results += "Creating new task: %s" % (task.full_name())

                if not item.BusyStatus:
                    task.booking_number = '#%s' % (subject)
                if start_time == end_time:
                    pass
                else:
                    start = self.new_event()
                    end = self.new_event()
                    if item.AllDayEvent:
                        start.time = start_time.replace(hour = 0,
                                                        minute = 0,
                                                        second = 0,
                                                        microsecond = 0)
                        end.time =  start_time.replace(hour = 23,
                                                       minute = 59,
                                                       second = 59,
                                                       microsecond = 0)
                    else:
                        start.time = start_time
                        end.time = end_time

                    # If this is not a recurring event use the entry id to find
                    # duplicates else test the task and start time.
                    #
                    new_activity = None
                    if not item.IsRecurring:
                        new_activity = find(lambda activity : activity.outlook_id == item.EntryID,
                                            self.activities)
                    if new_activity:
                        if config.get('outlook_precedence'):
                            new_activity.start_event = start
                            start.subsequent_activity.add(new_activity)
                            new_activity.end_event = end
                            end.prior_activity.add(new_activity)
                    else:
                        new_activity = find(lambda activity :
                                            (activity.task, activity.start_event.time) == (task, start_time),
                                            self.activities)
                    if new_activity:
                        results += "Skipping duplicate activity for task: %s\n" % (task.full_name())
                    else:
                        results += "Adding activity for task: %s\n" % (task.full_name())
                        new_activity = self.new_activity(task, start, end)
                        new_activity.outlook_id = item.EntryID
                        if item.IsRecurring:
                            task.recurring_activities.append(item.EntryID)
        if results:
            Message.show('Import Results', results, gtk.MESSAGE_INFO)

        if function:
            function()

    def save(self):
        """Writes the current session file to disk."""
        path, extension = os.path.splitext(config.get('data_dir') + config.get('session_file'))
        filename = "%s_%s%s" % (path, self.week_end.date().isoformat(), extension)
        self.to_xml(filename)
        if config.get('daily_backup'):
            directory = os.path.expanduser(config.get('data_dir') + '/backup')
            if not os.path.exists(directory):
                os.makedirs(directory)
            filename = os.path.expanduser(filename)
            backup = directory + '/' + config.get('session_file')
            path, extension = os.path.splitext(backup)
            day = datetime.date.today().strftime('%A')
            backup = "%s_%s%s" % (path, day, extension)
            shutil.copy(filename, backup)

    def new_activity(self, task, start_event, end_event=None):
        """Creates a new activity."""
        new_activity = Activity(self, task)
        self.activities.append(new_activity)
        new_activity.change_start(start_event)
        if end_event:
            new_activity.change_end(end_event)
        else:
            new_activity.end_event = Now(self)
            self.current_activity = new_activity
        return new_activity

    def new_event(self, prior_activity=None, subsequent_activity=None, label=None):
        """Creates a new event."""
        new_event = Event(self, prior_activity, subsequent_activity, label)
        self.events.append(new_event)
        return new_event

    def start_task(self, task):
        """Start the specified task and update the session file."""
        stop_event = None
        if task:
            if not task.is_runable():
                Message.show("Inactive task", "Starting inactive task %s" % (task.name), gtk.MESSAGE_WARNING)

            if not self.current_activity or task != self.current_activity.task:
                if task.state == 'Pending':
                    task.state ='Active'
                new_activity = self.new_activity(task, self.new_event())
                self.task_manager.task_activated(task)

            self.to_xml(config.get('data_dir') + config.get('session_file'))

            self.update_gui()
            Message.show("Activate task",
                         "Task %s activated. As activity %d." % (self.current_activity.task.full_name(),
                                                                 self.current_activity.get_id()),
                         gtk.MESSAGE_INFO)
        else:
            Message.show("Invalid Task!", "Activate called with null task.", gtk.MESSAGE_ERROR)

    def pause(self, new_event=None):
        """Pauses the current activity recording."""
        if self.current_activity:
            if not new_event:
                new_event = self.new_event()
            self.previous_task = self.current_activity.task
            self.current_activity = None
            self.to_xml(config.get('data_dir') + config.get('session_file'))
            self.update_gui()
            Message.show("Paused task", "Task %s paused." % (self.previous_task.full_name()), gtk.MESSAGE_INFO)
        else:
            Message.show("No active Task!", "Attempted pause while not running a task.", gtk.MESSAGE_WARNING)
        return new_event

    def resume(self):
        """Resumes a paused task."""
        if self.previous_task:
            self.start_task(self.previous_task)
        else:
            Message.show("No Previous Task!", "Attempted resume without a previous task.", gtk.MESSAGE_WARNING)

    def find_child_by_name(self, name):
        """Returns the child task by name."""
        return find(lambda task: task.name == name, self.children)

    def find_descendent(self, task, path):
        """Returns the descendent task from the path."""
        path = path.strip('/')
        first = path.find('/')
        sub_path = None
        if first > 0:
            descendent = find(lambda task: task.name == path[:first], task.children)
            sub_path = path[first:]
        else:
            descendent = find(lambda task: task.name == path, task.children)

        if descendent and sub_path:
            descendent = self.find_descendent(descendent, sub_path)

        return descendent

    def shift_event(self, event, timedelta):
        """Shifts the event and any following events by the given time."""
        event.time += timedelta
        self.shifted_events.append(event)
        for activity in event.subsequent_activity:
            if activity.end_event not in self.shifted_events:
                self.shift_event(activity.end_event, timedelta)
        for activity in event.prior_activity:
            if activity.start_event not in self.shifted_events:
                self.shift_event(activity.end_event, timedelta)

    def check_upcoming_events(self):
        """Returns a list of events that are due to occur within their reminder period."""
        now = datetime.datetime.now()
        if (self.current_activity and 
            self.current_activity.end_event and
            self.current_activity.end_event.time < now):
            self.new_activity(self.current_activity.task, self.current_activity.end_event)
            
        start_activity = [event.action for event in self.events if isinstance(event.action, Activity)]
        if len(start_activity) == 1 and now > start_activity[0].start_event.time:
            self.current_activity = start_activity[0]
            del start_activity[0].start_event.action

        delayed_events = [event for event in self.events if event.action == 'Delay']
        shifted_events = []
        for event in delayed_events:
            time_delta = now - event.time
            if time_delta > datetime.timedelta():
                self.shift_event(event, time_delta)

        events = [event for event in self.events if (event.time - now) < event.reminder and event.remind]
        if self.current_activity and self.current_activity.start_event in events:
            events.remove(self.current_activity.start_event)
        return events

    def show(self):
        """Display the times for current tasks."""
        for activity in self.activities:
            print (activity)

    def to_xml(self, filename):
        """Writes the current session data to file."""
        filename = os.path.expanduser(filename)

        document = xml.dom.minidom.getDOMImplementation().createDocument(None, None, None)

        root_node = document.createElement('session')

        if self.current_activity:
            active_task_node = document.createElement('current_activity')
            active_task_node.appendChild(document.createTextNode(str(self.current_activity.get_id())))
            root_node.appendChild(active_task_node)

        if self.previous_task:
            previous_task_node = document.createElement('previous_task')
            previous_task_node.appendChild(document.createTextNode(str(self.previous_task.get_id())))
            root_node.appendChild(previous_task_node)

        activities_node = document.createElement('activities')
        activities = [(activity.get_id(), activity) for activity in self.activities]
        activities.sort()
        for id, activity in activities:
            activities_node.appendChild(activity.to_xml(document))
            if activity not in activity.start_event.subsequent_activity:
                Message.show('Inconsistent State',
                             'Inconsistent state for activity %s\nActivity not listed in start event.' % (str(activity)),
                             gtk.MESSAGE_ERROR)
            if activity.end_event and activity not in activity.end_event.prior_activity:
                Message.show('Inconsistent State',
                             'Inconsistent state for activity %s\nActivity not listed in end event.' % (str(activity)),
                             gtk.MESSAGE_ERROR)
        root_node.appendChild(activities_node)

        events_node = document.createElement('events')
        events = [(event.get_id(), event) for event in self.events
                  if event.prior_activity or event.subsequent_activity or event.label]
        events.sort()
        for id, event in events:
            events_node.appendChild(event.to_xml(document))
        root_node.appendChild(events_node)

        document.appendChild(root_node)

        if not os.path.exists(os.path.dirname(filename)):
            os.makedirs(os.path.dirname(filename))
        file = open(filename, 'w')
        file.write(document.toprettyxml())
        file.close()

    def duration_from_xml(self, node):
        """Load the duration data from an XML node."""
        start_event = None
        end_event = None
        for child_node in node.childNodes:
            if child_node.tagName == 'task':
                task = int(child_node.firstChild.nodeValue)
            elif child_node.tagName == 'start_event':
                start_event = int(child_node.firstChild.nodeValue)
            elif child_node.tagName == 'end_event':
                end_event = int(child_node.firstChild.nodeValue)
        return (task, start_event, end_event)

    def from_xml(self, filename):
        """Load an xml session file."""
        result = False
        pattern = re.compile("\s*(<|>)\s*")
        filename = os.path.expanduser(filename)
        activities = []

        if os.path.exists(filename):
            file = open(filename)
            xml_string = file.read()
            file.close()

            # Remove pretty printing
            xml_string = pattern.sub(r'\1', xml_string)
            document = xml.dom.minidom.parseString(xml_string)
            current_activity = None
            previous_task = None
            for node in document.documentElement.childNodes:
                if node.tagName == 'current_activity':
                    current_activity = int(node.firstChild.nodeValue)
                elif node.tagName == 'previous_task':
                    previous_task = int(node.firstChild.nodeValue)
                elif node.tagName == 'task':
                    new_task = Task()
                    if new_task:
                        new_task.from_xml(node)
                        self.tasks.append(new_task)
                elif node.tagName == 'activity':
                    new_activity = Activity(self, None)
                    new_activity.from_xml(node)
                    activities.append(new_activity)
                    self.activities.append(new_activity)
                elif node.tagName == 'activities':
                    for child_node in node.childNodes:
                        new_activity = Activity(self, None)
                        new_activity.from_xml(child_node)
                        activities.append(new_activity)
                        self.activities.append(new_activity)
                elif node.tagName == 'events':
                    for child_node in node.childNodes:
                        new_event = self.new_event()
                        new_event.from_xml(child_node)
                else:
                    Message.show('Unexpected xml tag',
                                 'Session file %s has node with unexpected tag: %s.' % (filename, node.tagName),
                                 gtk.MESSAGE_INFO)

            self.resolve_ids(current_activity, previous_task)
            result = True

        return result

    def resolve_ids(self, current_activity, previous_task):
        """Replace integer indicies with references to object."""
        # Create indexes to avoid repeating searches
        activities = {}
        for activity in self.activities:
            activities[activity.get_id()] = activity

        events = {0:Now(self)}
        for event in self.events:
            events[event.get_id()] = event

        tasks = {}
        # Iterate over all tasks ignoring heirachy
        for task in expand_list(self.task_manager.tasks):
            tasks[task.get_id()] = task

        invalid_activities = []
        for activity in [item for item in self.activities]:
            try:
                activity.task = tasks[activity.task]
            except KeyError:
                Message.show('Invalid Task',
                             ('Task %d referenced by activity %d\n' % 
                              (activity.task, activity.get_id())) +
                             'is not in the task list!',
                             gtk.MESSAGE_WARNING)
                self.activities.remove(activity)
                continue
            if activity.start_event:
                if activity.start_event in events:
                    event = events[activity.start_event]
                    if activity.get_id() not in event.subsequent_activity:
                        Message.show('Inconsistent State',
                                     'Fixing inconsistent start event for activity %d (%s)' % (activity.get_id(),
                                                                                               activity.task.name),
                                     gtk.MESSAGE_WARNING)
                        event.subsequent_activity.add(activity.get_id())
                    activity.start_event = event
                else:
                    activity.start_event = None
                    Message.show('Invalid Activity', 'Activity %d has no start event!' % (activity.get_id()),
                                 gtk.MESSAGE_WARNING)
                    invalid_activities.append(activity)
            if activity.end_event:
                if activity.end_event in events:
                    event = events[activity.end_event]
                    if activity.get_id() not in event.prior_activity:
                        Message.show('Inconsistent State',
                                     'Fixing inconsistent end event for activity %d (%s)' % (activity.get_id(),
                                                                                             activity.task.name),
                                     gtk.MESSAGE_WARNING)
                        event.prior_activity.add(activity.get_id())
                    activity.end_event = event
                else:
                    activity.end_event = None
                    #invalid_activities.append(activity)
            elif not current_activity == activity.get_id():
                Message.show('Invalid Activity',
                        'Activity %d has no end event and is not the current event!' % (activity.get_id()),
                        gtk.MESSAGE_WARNING)
                invalid_activities.append(activity)
            else:
                activity.end_event = Now(self)

        events_list = [event for event in self.events]
        events_list.sort(key=operator.attrgetter('time'))
        for activity in invalid_activities:
            if activity.end_event and not activity.start_event:
                index = events_list.index(activity.end_event)
                if index > 0:
                    Message.show('Inconsistent State',
                                 'Moving invalid start event for activity %d (%s)' % (activity.get_id(),
                                                                                      activity.task.name),
                                 gtk.MESSAGE_WARNING)
                    activity.start_event = events_list[index - 1]
                    activity.start_event.subsequent_activity.add(activity.get_id())
                else:
                    Message.show('Inconsistent State',
                                 'Creating missing start event for activity %d (%s)' % (activity.get_id(),
                                                                                        activity.task.name),
                                 gtk.MESSAGE_WARNING)
                    activity.start_event = self.new_event()
                    activity.start_event.time = activity.end_event.time - datetime.timedelta(hour = 1)
            else:
                Message.show('Inconsistent State',
                             'removing activity %d (%s) invalid start and end' % (activity.get_id(),
                                                                                  activity.task.name),
                             gtk.MESSAGE_WARNING)
                activity.delete()

        for event in self.events:
            for id in [id for id in event.prior_activity]:
                event.prior_activity.remove(id)
                if id in activities:
                    if activities[id].end_event == event:
                        event.prior_activity.add(activities[id])
                    else:
                       Message.show('Inconsistent State',
                                    'Inconsistent reference to prior activity %d in event %s' % (id, str(event)),
                                    gtk.MESSAGE_WARNING)
                else:
                    Message.show('Inconsistent State',
                                 'Unresolved reference to prior activity %d in event %s' % (id, str(event)),
                                 gtk.MESSAGE_WARNING)

            for id in [id for id in event.subsequent_activity]:
                event.subsequent_activity.remove(id)
                if id in activities:
                    if activities[id].start_event == event:
                        event.subsequent_activity.add(activities[id])
                    else:
                        Message.show('Inconsistent State',
                                     'Inconsistent reference to subsequent activity %d in event %s' % (id, event),
                                     gtk.MESSAGE_WARNING)
                else:
                    Message.show('Inconsistent State',
                                 'Unresolved reference to sebsequent activity %d in event %s' % (id, event),
                                  gtk.MESSAGE_WARNING)

        self.normalise_data()
        self.current_activity = None
        if current_activity:
            if current_activity in activities:
                self._current = activities[current_activity]
            else:
                Message('Invalid Reference',
                        'Current activity index %d not valid!' % (current_activity),
                        gtk.MESSAGE_WARNING)
                
        if previous_task:
            self.previous_task = tasks[previous_task]


    def normalise_data(self):
        """Ensure that each preceding day is complete and that a maximum of one task is active."""
        day_start = datetime.datetime.today().replace(hour = int(config.get('day_start')[:2]),
                                                      minute=int(config.get('day_start')[3:]),
                                                      second=0,
                                                      microsecond=0)
        day_end = datetime.datetime.today().replace(hour = int(config.get('day_end')[:2]),
                                                    minute=int(config.get('day_end')[3:]),
                                                    second=0,
                                                    microsecond=0)

        events = [event for event in self.events]
        for activity in self.activities:
            if not (config.get('allow_unfinished') or activity.end_event):
                events.sort(key=operator.attrgetter('time'))
                index = events.index(activity.start_event)
                if index == len(events) - 1 or events[index + 1].time.date() > activity.start_event.time.date():
                    Message.show("Adding end event",
                                 "Adding end event for activity %d (%s)" % (activity.get_id(), activity.task.name),
                                 gtk.MESSAGE_WARNING)
                    activity.end_event = self.new_event(activity)
                    events.append(event)
                    if activity.start_event.time.date() < datetime.datetime.today().date() or \
                        activity.start_event.time > datetime.datetime.now():
                        if activity.start_event.time.time() > day_end.time():
                            activity.end_event.time = activity.start_event.time.replace(hour = 23,
                                                                                        minute = 59,
                                                                                        second = 59,
                                                                                        microsecond = 0)
                        else:
                            date = activity.start_event.time.date()
                            activity.end_event.time = datetime.datetime.combine(date, day_end.time())
                else:
                    Message.show("Setting end event",
                                 "Setting %d (%s) end event to %d" % (activity.get_id(),
                                                                      activity.task.name,
                                                                      events[index + 1].get_id()),
                                 gtk.MESSAGE_WARNING)
                    activity.end_event = events[index + 1]
                    activity.end_event.prior_activity.add(activity)

            if not config.get('allow_overnight'):
                # Check for activity over day boundary
                if activity.end_event and activity.start_event.time.day != activity.end_event.time.day:
                    Message.show("Splitting activity",
                                 "Splitting overnight activity %d (%s)" % (activity.get_id(), activity.task.name),
                                 gtk.MESSAGE_WARNING)

                    new_end = self.new_event()
                    new_end.time = activity.start_event.time.replace(hour = int(config.get('day_end')[:2]),
                                                                     minute = int(config.get('day_end')[3:]),
                                                                     second = 0,
                                                                     microsecond = 0)
                    if new_end.time < activity.start_event.time:
                        new_end.time.replace(hour = 23, minute = 59, second = 59, microsecond = 0)
                    new_start = self.new_event()
                    new_start.time = activity.end_event.time.replace(hour = int(config.get('day_start')[:2]),
                                                                     minute = int(config.get('day_start')[3:]),
                                                                     second = 0,
                                                                     microsecond = 0)
                    if new_start.time > activity.end_event.time:
                        new_end.time.replace(hour = 0, minute = 0, second = 0, microsecond = 0)
                    if activity.end_event.get_id():
                        new_activity = self.new_activity(activity.task,
                                                         new_start,
                                                         activity.end_event)
                        activity.change_end(new_end)
                    else:
                        new_activity = self.new_activity(activity.task,
                                                         new_start)
                        activity.end_event = new_end
                        new_end.prior_activity.add(activity)

        if config.get('join_activities'):
            events.sort(key=operator.attrgetter('time'))
            for event in events:
                for activity in [prior for prior in event.prior_activity]:
                    delete_activities = set()
                    for other_activity in event.subsequent_activity:
                        if activity == other_activity:
                            Message.show('Invalid State',
                                         'Event %d in invalid state activity %d in prior and subsequent list' % (event.get_id(),
                                                                                                                 activity.get_id()),
                                         gtk.MESSAGE_ERROR)
                            continue
                        if activity.task == other_activity.task:
                            if other_activity.end_event.time > activity.end_event.time:
                                current_end = activity.end_event
                                activity.end_event = other_activity.end_event
                                activity.end_event.prior_activity.add(activity)
                                current_end.prior_activity.remove(activity)
                            delete_activities.add(other_activity)
                    for activity in delete_activities:
                        activity.delete()

        for activity in self.activities:
            if activity.end_event and activity.start_event.time > activity.end_event.time:
                activity.start_event, activity.end_event = activity.end_event, activity.start_event

        for event in [event for event in self.events]:
            event.try_delete()

class Task_Dialog(gtk.Dialog, Managed_Dialog):
    """Task editing dialog."""

    __gsignals__ = {'changed' : (gobject.SIGNAL_RUN_LAST, gobject.TYPE_NONE, (gobject.TYPE_INT,))}

    def __init__(self, manager, session, task=None, parent=None, start=False):
        gtk.Dialog.__init__(self, "New Task", None, 0,
                            (gtk.STOCK_CANCEL, gtk.RESPONSE_REJECT,
                             gtk.STOCK_OK, gtk.RESPONSE_ACCEPT))
        name = 'task_%d_window' % (task.get_id()) if task else 'task_window'
        Managed_Dialog.__init__(self, name, persistent_state=False) 

        icons = []
        for image_file in glob.glob(config.get('image_root') + "*x*/actions/view-pim-tasks.png"):
            icons.append(gtk.gdk.pixbuf_new_from_file(image_file))
        self.set_icon_list(*icons)

        self.task_manager = manager
        self.session = session
        if task:
            self.task = task
            self.new = False
        else:
            self.task = Task(parent)
            self.new = True

        self.connect('response', self.response_cb)
        self.connect('map-event', self.mapped_cb)
        path_buttons = gtk.HBox()
        path_buttons.pack_start(gtk.Label('Parent:'), False, False)
        self.update_path_widget(path_buttons)

        self.vbox.pack_start(path_buttons, False, False)
        task_box = gtk.Table(2, 7)
        task_label = gtk.Label("Task Name:")
        task_label.set_alignment(1.0, 0.5)
        self.task_name = gtk.Entry()
        booking_no_label = gtk.Label("Booking Number:")
        booking_no_label.set_alignment(1.0, 0.5)
        self.booking_no = gtk.Entry()

        task_box.attach(task_label, 0, 1, 0, 1, gtk.FILL, gtk.FILL, 3, 3)
        task_box.attach(self.task_name, 1, 2, 0, 1, gtk.FILL | gtk.EXPAND, gtk.FILL, 3, 3)
        task_box.attach(booking_no_label, 0, 1, 1, 2, gtk.FILL, gtk.FILL, 3, 3)
        task_box.attach(self.booking_no, 1, 2, 1, 2, gtk.FILL | gtk.EXPAND, gtk.FILL, 3, 3)

        self.vbox.pack_start(task_box, False, False)

        task_box = gtk.Table(4, 7)
        start_label = gtk.Label("Start Date:")
        start_label.set_alignment(1.0, 0.5)
        self.start_date = Day_Selector()
        end_label = gtk.Label("End Date:")
        end_label.set_alignment(1.0, 0.5)
        self.end_date = Day_Selector()

        allocated_label = gtk.Label("Allocated:")
        allocated_label.set_alignment(1.0, 0.5)
        self.allocated = gtk.Entry()
        estimated_label = gtk.Label("Estimated:")
        estimated_label.set_alignment(1.0, 0.5)
        self.estimated = gtk.Entry()

        priority_label = gtk.Label("Priority:")
        priority_label.set_alignment(1.0, 0.5)
        self.priority = gtk.Entry()
        state_label = gtk.Label("State:")
        state_label.set_alignment(1.0, 0.5)
        self.task_state = gtk.combo_box_new_text()
        [self.task_state.append_text(state) for state in Task.states if state != 'Deleted' and state != 'New']

        task_box.attach(start_label, 0, 1, 3, 4, gtk.FILL, gtk.FILL, 3, 1)
        task_box.attach(self.start_date, 1, 2, 3, 4, gtk.FILL | gtk.EXPAND, gtk.FILL, 3, 1)
        task_box.attach(end_label, 2, 3, 3, 4, gtk.FILL, gtk.FILL, 3, 1)
        task_box.attach(self.end_date, 3, 4, 3, 4, gtk.FILL | gtk.EXPAND, gtk.FILL, 3, 1)

        task_box.attach(allocated_label, 0, 1, 4, 5, gtk.FILL, gtk.FILL, 3, 3)
        task_box.attach(self.allocated, 1, 2, 4, 5, gtk.FILL | gtk.EXPAND, gtk.FILL, 3, 3)
        task_box.attach(estimated_label, 2, 3, 4, 5, gtk.FILL, gtk.FILL, 3, 3)
        task_box.attach(self.estimated, 3, 4, 4, 5, gtk.FILL | gtk.EXPAND, gtk.FILL, 3, 3)

        task_box.attach(priority_label, 0, 1, 5, 6, gtk.FILL, gtk.FILL, 3, 3)
        task_box.attach(self.priority, 1, 2, 5, 6, gtk.FILL | gtk.EXPAND, gtk.FILL, 3, 3)
        task_box.attach(state_label, 2, 3, 5, 6, gtk.FILL, gtk.FILL, 3, 3)
        task_box.attach(self.task_state, 3, 4, 5, 6, gtk.FILL | gtk.EXPAND, gtk.FILL, 3, 3)

        self.vbox.pack_start(task_box, False, False)

        notes_label = gtk.Label("Notes:")
        notes_label.set_alignment(0.0, 0.0)
        self.notes = Notes()
        self.notes.view.get_buffer().register_serialize_tagset()
        self.notes.view.get_buffer().register_deserialize_tagset()
        self.notes.grab_focus()

        self.vbox.pack_start(notes_label, False, False)
        self.vbox.pack_start(self.notes)

        if self.session and self.session.is_current():
            if not self.new:
                if (task.is_runable() and (not self.session.current_activity or
                    task != self.session.current_activity.task)):
                    self.activate = gtk.Button('Start Now')
                    self.activate.connect('clicked', self.start_now_cb)
                    self.action_area.pack_start(self.activate, False, False)
                    self.action_area.reorder_child(self.activate, 0)
            else:
                self.activate = gtk.CheckButton("Start Task")
                self.activate.set_active(start)
                self.action_area.pack_start(self.activate, False, False)
                self.action_area.reorder_child(self.activate, 0)

        self.set_title(self.task.name)
        self.populate_widgets()
        self.show_all()

    def mapped_cb(self, widget, event):
        """Register the window with the drag and drop handler."""
        if self.window:
            self.drop = dragndrop.Drop_Target(self.notes)

    def new_dialog_cb(self, button, task=None):
        """Display a task dialog for the selected task."""
        Task_Dialog(self.session.task_manager, self.session, task) 
        
    def update_path_widget(self, container):
        """Updates the widget representing the path to the current task."""
        for button in container.get_children()[1:]:
            container.remove(button)
        parent = self.task.parent
        parent_buttons = []
        while parent:
            parent_button = gtk.Button(parent.name)
            parent_button.set_relief(gtk.RELIEF_NONE)
            parent_button.connect('clicked', self.change_parent_cb, parent)
            parent_buttons.insert(0, parent_button)
            parent = parent.parent

        for button in parent_buttons:
            container.pack_start(button, False, False)
            container.pack_start(gtk.Label('/'), False, False)

        if self.task.parent and self.task.parent.children:
            add_button = gtk.Button('...')
            add_button.set_relief(gtk.RELIEF_NONE)
            add_button.connect('clicked', self.popup_task_menu)
            container.pack_start(add_button, False, False)
        container.show_all()

    def change_parent_cb(self, button, new_parent):
        """Changes the parent to the selected parent."""
        self.task_manager.reparent(self.task, new_parent)

        path_buttons = self.vbox.get_children()[0]
        self.update_path_widget(path_buttons)

    def popup_task_menu(self, button):
        """Popup a menu of tasks."""
        if self.task.parent:
            menu = [(task.name, self.change_parent_cb, task)
                    for task in self.task.parent.children
                    if task != self.task]
        else:
            menu = [(task.name, self.change_parent_cb, task)
                    for task in self.session.task_manager.tasks
                    if task != self.task]
        menu = build_menu(menu)
        menu.popup(None, None, None, 1, gtk.get_current_event_time())

    def start_now_cb(self, button):
        """Responds to a click on the start now button."""
        if self.task and self.session and self.session.is_current():
            self.session.start_task(self.task)
            self.action_area.remove(self.activate)

    def response_cb(self, dialog, response_id):
        """Process the user response."""
        if response_id == gtk.RESPONSE_ACCEPT:
            start_task = False
            if self.session and self.new:
                start_task = self.activate.get_active()
            name = self.task_name.get_text()
            booking_no = self.booking_no.get_text()
            notes = self.notes.parse_buffer()
            self.task.name = name
            self.task.booking_number = booking_no
            self.task.start_date = self.start_date.date
            self.task.end_date = self.end_date.date
            if self.allocated.get_text():
                self.task.allocated_time = datetime.timedelta(hours=float(self.allocated.get_text()))
            else:
                self.task.allocated_time = datetime.timedelta()
            if self.estimated.get_text():
                self.task.estimated_time = datetime.timedelta(hours=float(self.estimated.get_text()))
            else:
                self.task.estimated_time = datetime.timedelta()
            if self.priority.get_text():
                self.task.priority = int(self.priority.get_text())
            else:
                self.task.priority = 0
            self.task.state = self.task_state.get_active() + 2
            self.task.notes = notes
            if self.new:
                self.task.state = 'Pending'
                self.task_manager.reparent(self.task, self.task.parent)

            if self.session and start_task:
                self.session.start_task(self.task)
            self.task_manager.write_task_list()
        else:
            if self.new:
                if self.task.parent and self.task in self.task.parent.children:
                    self.task.parent.children.remove(self.task)
                elif self.task in self.task_manager.tasks:
                    self.task_manager.tasks.remove(self.task)
               
        self.destroy()

    def populate_widgets(self):
        """Sets the widget values to the data from the relevent task."""
        self.task_name.set_text(self.task.name)
        if self.task.booking_number:
            self.booking_no.set_text(str(self.task.booking_number))
        self.start_date.set_date(self.task.start_date)
        self.end_date.set_date(self.task.end_date)
        self.allocated.set_text('%02.2f' % (get_decimal_hours(self.task.allocated_time)))
        self.estimated.set_text('%02.2f' % (get_decimal_hours(self.task.estimated_time)))
        self.priority.set_text(str(self.task.priority))
        self.task_state.set_active(self.task._state - 2)
        self.notes.display_notes(self.task.notes)

    def set_defaults(self):
        """Sets the widget values to the data from the relevent task."""
        self.allocated.set_text('0.0')
        self.estimated.set_text('0.0')
        self.priority.set_text('0')
        self.task_state.set_active(0)

    def open_link(self, link):
        """Displays GUI ellement associated with linked item."""
        task = self.session.task_manager.find_by_path(link)
        if task:
            Task_Dialog(self.session.task_manager, self.session, task)

class Options_Dialog(gtk.Dialog, Managed_Dialog):
    """Options GUI class"""

    def __init__(self):
        gtk.Dialog.__init__(self, "Time tracker options", None, 0,
                            (gtk.STOCK_CANCEL, gtk.RESPONSE_CANCEL,
                             gtk.STOCK_OK, gtk.RESPONSE_OK,
                             gtk.STOCK_APPLY, gtk.RESPONSE_APPLY))
        Managed_Dialog.__init__(self, 'options_dialog')
        icons = []
        for image_file in glob.glob(config.get('image_root') + "*x*/apps/preferences-system.png"):
            icons.append(gtk.gdk.pixbuf_new_from_file(image_file))
        self.set_icon_list(*icons)

        self.connect('response', self.response_cb)
        packing = gtk.HBox()
        startup_options = gtk.Frame("Task To Activate On Startup")
        startup_packing = gtk.VBox()
        self.no_task = gtk.RadioButton(None, "None")
        self.ask = gtk.RadioButton(self.no_task, "Ask")
        self.previous_task = gtk.RadioButton(self.no_task, "Previous")
        startup_packing.pack_start(self.no_task)
        startup_packing.pack_start(self.ask)
        startup_packing.pack_start(self.previous_task)
        startup_options.add(startup_packing)
        packing.pack_start(startup_options)

        display_options = gtk.Frame("Task Display Format")
        display_packing = gtk.VBox()
        self.hierarchy = gtk.RadioButton(None, "Heirarchy")
        self.flat = gtk.RadioButton(self.hierarchy, "Flat")
        display_packing.pack_start(self.hierarchy)
        display_packing.pack_start(self.flat)
        display_options.add(display_packing)
        packing.pack_start(display_options)

        information_options = gtk.Frame("Information")
        information_packing = gtk.VBox()
        self.all = gtk.RadioButton(None, "All")
        self.warnings = gtk.RadioButton(self.all, "Warnings and Errors")
        self.errors = gtk.RadioButton(self.all, "Errors only")
        information_packing.pack_start(self.all)
        information_packing.pack_start(self.warnings)
        information_packing.pack_start(self.errors)
        information_options.add(information_packing)
        packing.pack_start(information_options)
        self.vbox.pack_start(packing)
        self.set_state()
        self.show_all()

    def set_state(self):
        if not config.get('startup_task'):
            self.no_task.set_active(True)
        elif config.get('startup_task') == 'ask':
            self.ask.set_active(True)
        elif config.get('startup_task') == 'previous':
            self.previous_task.set_active(True)

        if config.get('hierarchical_list'):
            self.hierarchy.set_active(True)
        else:
            self.flat.set_active(True)

        if config.get('verbosity') == 7:
            self.all.set_active(True)
        elif config.get('verbosity') == 3:
            self.warnings.set_active(True)
        elif config.get('verbosity') == 1:
            self.errors.set_active(True)

    def get_state(self):
        if self.no_task.get_active():
            config['startup_task'] = None
        elif self.ask.get_active():
            config['startup_task'] = 'ask'
        elif self.previous_task.get_active():
            config['startup_task'] = 'previous'

        if self.hierarchy.get_active():
            config['hierarchical_list'] = True
        elif self.flat.get_active():
            config['hierarchical_list'] = False

        if self.all.get_active():
            config['verbosity'] = 7
        elif self.warnings.get_active():
            config['verbosity'] = 3
        elif self.errors.get_active():
            config['verbosity'] = 1

    def response_cb(self, dialog, response_id):
        self.destroy()
        if response_id != gtk.RESPONSE_CANCEL:
            self.get_state()
        if response_id == gtk.RESPONSE_OK:
            config.save()

class Notepad(gtk.Window, Managed_Window):
    """Widget to create general notes."""
    
    __gsignals__ = {'changed' : (gobject.SIGNAL_RUN_LAST, gobject.TYPE_NONE, ())}

    __instance = None

    @classmethod
    def instance(cls, session):
        """Return the notepad instance if set else create a new instance."""
        if not cls.__instance:
            cls.__instance = cls(session)
        cls.__instance.present()
        return cls.__instance

    def __init__(self, session):
        """Create and display the notepad."""
        gtk.Window.__init__(self, gtk.WINDOW_TOPLEVEL)
        Managed_Window.__init__(self, 'notepad')
        
        icons = []
        for image_file in glob.glob(config['image_root'] + '*x*/apps/accessories-text-editor.png'):
            icons.append(gtk.gdk.pixbuf_new_from_file(image_file))
        self.set_icon_list(*icons)

        config['notepad']['show'] = True
        self.session = session
        self.connect('map-event', self.mapped_cb)
        self.connect("delete_event", self.delete_cb)

        self.add(Notes())
        self.child.view.set_wrap_mode(gtk.WRAP_NONE)
        self.child.display_notes(self.session.task_manager.notes)
        self.child.connect('key-release-event', self.key_release_cb)

        self.show_all()

    def delete_cb(self, window, event):
        """Set the instance state to indicate the window is not present."""
        self.__class__.__instance = None
        self.session.task_manager.notes = self.child.parse_buffer()
        self.session.task_manager.write_task_list()

    def mapped_cb(self, widget, event):
        """Register the window with the drag and drop handler."""
        if self.window:
            self.drop = dragndrop.Drop_Target(self.get_child())

    def key_release_cb(self, widget, event):
        """Processes keys specific to the notes widget."""
        if (event.state & gtk.gdk.CONTROL_MASK and
            event.keyval == gtk.keysyms.s):
            self.session.task_manager.notes = widget.parse_buffer()
            self.session.task_manager.write_task_list()
            return True

    def destroy(self):
        """Save the buffer contents before closing the window."""
        self.session.task_manager.notes = self.child.parse_buffer()
        self.session.task_manager.write_task_list()
        Managed_Window.destroy(self)

    def update(self):
        """The task / activity data has changed."""
        pass

    def redraw(self):
        """The task / activity data has changed."""
        pass

    def open_link(self, link):
        """Displays GUI ellement associated with linked item."""
        task = self.session.task_manager.find_by_path(link)
        if task:
            Task_Dialog(self.session.task_manager, self.session, task)

class Task_List(gtk.Window, Managed_Window):
    """Customisable list of tasks."""

    sort_options = {'priority' : 'oooPriority',
                    'end_date' : 'oooEnd Date',
                    'start_date' : 'oooStart Date',
                    'name' : 'oooName',
                    'recent' : 'oooMost Recent'}

    __gsignals__ = {'changed' : (gobject.SIGNAL_RUN_LAST, gobject.TYPE_NONE, ()),
                    'timesheet-request' : (gobject.SIGNAL_RUN_LAST, gobject.TYPE_NONE, ()),
                    'find-task-request' : (gobject.SIGNAL_RUN_LAST, gobject.TYPE_NONE, ()),
                    'notepad-request' : (gobject.SIGNAL_RUN_LAST, gobject.TYPE_NONE, ()),
                    'show-task-selection' : (gobject.SIGNAL_RUN_LAST, gobject.TYPE_NONE, ()),
                    'exit-request' : (gobject.SIGNAL_RUN_LAST, gobject.TYPE_NONE, ())}

    def __init__(self, session):
        """Generate the task list and display the window."""
        gtk.Window.__init__(self, gtk.WINDOW_POPUP)
        Managed_Window.__init__(self, 'todo')

        config['todo']['show'] = True

        # Only want focus when mouse is over the window
        self.set_focus_on_map(False)

        icons = []
        for image_file in glob.glob(config.get('image_root') + "*x*/apps/preferences-system-time.png"):
            icons.append(gtk.gdk.pixbuf_new_from_file(image_file))
        self.set_icon_list(*icons)

        self.set_type_hint(gtk.gdk.WINDOW_TYPE_HINT_MENU)
        self.set_property('skip_pager_hint', True)
        self.set_property('skip_taskbar_hint', True)
        self.set_decorated(False)
        self.opacity = float(config['todo'].get('opacity'))
        self.session = session
        if config['todo'].get('sort_key') != 'recent':
            self.key_list = [config['todo'].get('sort_key')]
        else:
            self.key_list = ['priority']
        self.show_full_name = config['todo'].get('full_name')
        self.list_size = config['todo'].get('size')
        self.keep_above = config['todo'].get('keep_above')
        self.set_keep_above(self.keep_above)
        self.requested_position = config['todo'].get('position')
        self.coords = config['todo'].get('x'), config['todo'].get('y')
        self.current_first = config['todo'].get('current_first')
        self.show_controls = config['todo'].get('show_buttons')
        self.move_window = False
        sort = [(Task_List.sort_options[key], self.sort_key_cb, key) for key in Task_List.sort_options.keys()]
        menu_definition = [('oxoCurrent Task First', self.current_first_cb),
                           ('>>>Sort by', sort),
                           ('>>>Tasks Displayed', [('ooo1', self.task_count_cb, 1),
                                                   ('ooo3', self.task_count_cb, 3),
                                                   ('ooo5', self.task_count_cb, 5),
                                                   ('ooo10', self.task_count_cb, 10),
                                                   ('ooo20', self.task_count_cb, 20),
                                                   ('oooAll', self.task_count_cb, 0)]),
                           ('oxoDisplay Full Name', self.full_name_cb),
                           ('---', '---'),
                           ('oxoKeep On Top', self.keep_above_cb),
                           ('>>>Position', [('oooTop Left', self.set_position_cb, 1),
                                            ('oooTop Right', self.set_position_cb, 2),
                                            ('oooBottom Left', self.set_position_cb, 3),
                                            ('oooBottom Right', self.set_position_cb, 4),
                                            ('oooCentre', self.set_position_cb, 5)]),
                           ('oxoShow Controls', self.show_controls_cb),
                           ('Hide List', self.close_cb)]

        if self.current_first:
            menu_definition[0] = ('oxx' + menu_definition[0][0][3:], menu_definition[0][1])
        index = Task_List.sort_options.keys().index(config['todo']['sort_key'])
        menu_definition[1][1][index] = ('oox' + menu_definition[1][1][index][0][3:],
                                        menu_definition[1][1][index][1],
                                        menu_definition[1][1][index][2])
        index = [1, 3, 5, 10, 20, 0].index(self.list_size)
        menu_definition[2][1][index] = ('oox' + menu_definition[2][1][index][0][3:],
                                        menu_definition[2][1][index][1],
                                        menu_definition[2][1][index][2])
        if self.show_full_name:
            menu_definition[3] = ('oxx' + menu_definition[3][0][3:], menu_definition[3][1])
        if self.keep_above:
            menu_definition[5] = ('oxx' + menu_definition[5][0][3:], menu_definition[5][1])
        if self.show_controls:
            menu_definition[7] = ('oxx' + menu_definition[7][0][3:], menu_definition[7][1])
        index = self.requested_position - 1
        if index > 0 and index < 6:
            menu_definition[6][1][index] = ('oox' + menu_definition[6][1][index][0][3:],
                                            menu_definition[6][1][index][1])

        self.menu = build_menu(menu_definition)

        self.list = gtk.ListStore(object, gobject.TYPE_STRING)
        self.view = gtk.TreeView(self.list)

        self.connect('popup-menu', self.context_menu_cb)
        self.connect('enter-notify-event', self.mouse_enter_cb)
        self.connect('leave-notify-event', self.mouse_leave_cb)
        self.connect('map-event', self.mapped_cb)
        self.view.connect('motion-notify-event', self.motion_cb)
        self.view.connect('button-press-event', self.button_cb)
        self.view.connect('button-release-event', self.button_release_cb)
        self.view.connect('row_activated', self.row_activated_cb)
        self.key_handler = self.view.connect('key-release-event', self.key_release_cb)
        self.change_handler = self.list.connect_after('row-changed', self.row_changed_cb)

        text_renderer = gtk.CellRendererText()
        tree_column = gtk.TreeViewColumn(None, text_renderer, markup = 1)
        tree_column.set_sizing(gtk.TREE_VIEW_COLUMN_AUTOSIZE)
        self.view.set_headers_visible(False)
        if self.key_list[0] == 'priority':
            self.view.set_reorderable(True)
        else:
            self.view.set_reorderable(False)
        self.view.append_column(tree_column)

        self.allow_update = True
        self.update()
        self.set_opacity(self.opacity / 100.0)

        box = gtk.VBox()
        self.controls = gtk.HBox()
        self.create_buttons(self.controls)
        box.pack_start(self.controls, False, False)
        box.pack_start(self.view)
        self.add(box)
        self.move(self.coords[0], self.coords[1])
        if len(self.list) > 0:
            self.show_all()
        if not self.show_controls:
            self.controls.hide()

    def create_buttons(self, parent):
        """Create the buttons for the top line of the widget."""
        config['todo']['buttons'] = config['image_root'] + '16x16/'
        change_task = gtk.Button()
        filename = config['todo']['buttons'] + 'actions/view-pim-tasks.png'
        if os.path.exists(filename):
            change_task.set_image(gtk.image_new_from_file(filename))
            change_task.set_relief(gtk.RELIEF_NONE)
            change_task.set_tooltip_text('Change the current task')
            change_task.connect('clicked', self.task_change_cb)
            parent.pack_start(change_task, False, False)

        timesheet = gtk.Button()
        filename = config['todo']['buttons'] + 'apps/preferences-system-time.png'
        if os.path.exists(filename):
            timesheet.set_image(gtk.image_new_from_file(filename))
            timesheet.set_relief(gtk.RELIEF_NONE)
            timesheet.set_tooltip_text('Display the timesheet window')
            timesheet.connect('clicked', self.timesheet_cb)
            parent.pack_start(timesheet, False, False)

        find = gtk.Button()
        filename = config['todo']['buttons'] + 'actions/find.png'
        if os.path.exists(filename):
            find.set_image(gtk.image_new_from_file(filename))
            find.set_relief(gtk.RELIEF_NONE)
            find.set_tooltip_text('Display the find task widget')
            find.connect('clicked', self.find_cb)
            parent.pack_start(find, False, False)

        notes = gtk.Button()
        filename = config['todo']['buttons'] + 'actions/edit.png'
        if os.path.exists(filename):
            notes.set_image(gtk.image_new_from_file(filename))
            notes.set_relief(gtk.RELIEF_NONE)
            notes.set_tooltip_text('Display the notes window')
            notes.connect('clicked', self.notepad_cb)
            parent.pack_start(notes, False, False)

        opacity = gtk.ScaleButton(gtk.ICON_SIZE_SMALL_TOOLBAR, 0, 100, 2, ('show-menu', 'show-menu'))
        #filename = config['todo']['buttons'] + 'actions/show-menu.png'
        #if os.path.exists(filename):
        #    opacity.set_image(gtk.image_new_from_file(filename))
        opacity.set_relief(gtk.RELIEF_NONE)
        opacity.set_tooltip_text('Change the window opacity')
        opacity.set_value(self.opacity)
        opacity.connect('value-changed', self.opacity_changed_cb)
        parent.pack_start(opacity, False, False)

        exit = gtk.Button()
        filename = config['todo']['buttons'] + 'actions/application-exit.png'
        if os.path.exists(filename):
            exit.set_image(gtk.image_new_from_file(filename))
            exit.set_relief(gtk.RELIEF_NONE)
            exit.set_tooltip_text('Exit the timesheet application')
            exit.connect('clicked', self.exit_cb)
            parent.pack_end(exit, False, False)

        hide = gtk.Button()
        filename = config['todo']['buttons'] + 'actions/fade.png'
        if os.path.exists(filename):
            hide.set_image(gtk.image_new_from_file(filename))
            hide.set_relief(gtk.RELIEF_NONE)
            hide.set_tooltip_text('Temporarily hide the task list')
            hide.connect('clicked', self.hide_cb)
            parent.pack_end(hide, False, False)

    def timesheet_cb(self, button):
        """Displays the timesheet widget."""
        self.emit('timesheet-request')

    def find_cb(self, button):
        """Displays the find task widget."""
        self.emit('find-task-request')

    def notepad_cb(self, button):
        """Displays the notepad widget."""
        self.emit('notepad-request')

    def task_change_cb(self, button):
        """Displays the task menu."""
        self.emit('show-task-selection')

    def exit_cb(self, button):
        """Emit an exit request."""
        self.emit('exit-request')

    def hide_cb(self, button):
        """Temporarily hide the window."""
        self.allow_updates = False
        self.hide()
        self.timeout = gobject.timeout_add(config['todo']['timeout'], self.redisplay_cb)

    def opacity_changed_cb(self, scalebutton, value):
        """Handles a change to the opacity slider."""
        self.opacity = value
        self.set_opacity(self.opacity / 100.0)
        config['todo']['opacitiy'] = int(self.opacity + 0.5)

    def current_first_cb(self, item):
        """Toggles the current first flag."""
        self.current_first = not self.current_first
        self.update()
        config['todo']['current_first'] = self.current_first

    def mapped_cb(self, widget, event):
        """Update the window position and set the keep above state."""
        self.update()
        self.set_keep_above(self.keep_above)

    def mouse_enter_cb(self, widget, event):
        """Grab the focus to allow quick hiding."""
        # Trying to do decent user interaction on WindowsXP is hopelessly
        # broken so don't bother!
        pass
        #gtk.gdk.keyboard_grab(widget.window)

    def mouse_leave_cb(self, widget, event):
        """Ungrab the focus."""
        # Trying to do decent user interaction on WindowsXP is hopelessly
        # broken so don't bother!
        pass
        #gtk.gdk.keyboard_ungrab()

    def full_name_cb(self, item):
        """Toggles the display of full task names."""
        self.show_full_name = not self.show_full_name
        self.update()
        config['todo']['full_name'] = self.show_full_name

    def keep_above_cb(self, item):
        """Toggles the current first flag."""
        self.keep_above = not self.keep_above
        self.set_keep_above(self.keep_above)
        config['todo']['keep_above'] = self.keep_above

    def show_controls_cb(self, item):
        """Toggles the current first flag."""
        self.show_controls = not self.show_controls
        if self.show_controls:
            self.controls.show_all()
        else:
            self.controls.hide()
        config['todo']['show_buttons'] = self.show_controls
        self.update()

    def set_position_cb(self, item, position):
        """Changes the position of the window."""
        self.requested_position = position
        self.update(False)
        config['todo']['position'] = position

    def task_count_cb(self, item, count):
        """Toggles the display of full task names."""
        self.list_size = count
        self.update()
        config['todo']['size'] = count

    def sort_key_cb(self, item, attribute):
        """Changes the primary sort key."""
        config['sort_key'] = attribute
        if attribute == 'priority':
            self.view.set_reorderable(True)
        else:
            self.view.set_reorderable(False)
        self.update()

    def button_cb(self, widget, event):
        """If right button display context menu."""
        if event.type == gtk.gdk.BUTTON_PRESS and event.button == 1:
            self.allow_update = False
            if event.state & gtk.gdk.MOD1_MASK:
                self.move_window = True
                self.coords = event.x_root, event.y_root

        if event.type == gtk.gdk.BUTTON_PRESS and event.button == 3:
            self.context_menu_cb(widget)

        if event.type == gtk.gdk.BUTTON_PRESS and event.button == 4:
            self.opacity -= 10
            self.opacity = max(self.opacity, 0)
            self.set_opacity(self.opacity / 100.0)

        if event.type == gtk.gdk.BUTTON_PRESS and event.button == 5:
            self.opacity += 10
            self.opacity = min(self.opacity, 100)
            self.set_opacity(self.opacity / 100.0)

    def button_release_cb(self, widget, event):
        """Sets the flag to allow updates once the mouse button is released."""
        self.allow_update = True
        if self.move_window:
            config['todo']['x'], config['todo']['y'] = self.get_position()
            self.update()
        self.move_window = False

    def motion_cb(self, widget, event):
        """Handles dragging the todo list around."""
        if self.move_window:
            start = self.get_position()
            diff = event.x_root - self.coords[0], event.y_root - self.coords[1]
            if diff[0] and diff[1]:
                self.requested_position = 0
                self.coords = int(start[0] + diff[0]), int(start[1] + diff[1])
                self.move(self.coords[0], self.coords[1])
                config['todo']['position'] = self.requested_position
            return True

    def context_menu_cb(self, widget):
        """Displays the context menu for the task tree."""
        self.menu.popup(None, None, None, 1, gtk.get_current_event_time())

    def row_activated_cb(self, view, path, view_column):
        """Display the task dialog for the selected row."""
        iter = self.list.get_iter(path)
        task = self.list.get_value(iter, 0)()
        Task_Dialog(self.session.task_manager, self.session, task)

    def row_changed_cb(self, model, path, iter):
        """Adjusts the task priority based on dropped position."""
        moved_task = model.get_value(iter, 0)()

        task_list = [task for task
                     in expand_list(self.session.task_manager.tasks)
                     if task.is_runable()]
        task_list.remove(moved_task)
        task_list.sort(key=operator.attrgetter('priority'))
        task_list.reverse()

        next = model.iter_next(iter)
        if next:
            lower_task = model.get_value(next, 0)()
            new_index = task_list.index(lower_task)
        else:
            new_index = len(self.list) - 2

        if new_index == 0:
            priority = task_list[0].priority + 1
        elif new_index == len(task_list):
            priority = task_list[-1].priority - 1
        else:
            higher_priority = task_list[new_index - 1].priority
            lower_priority = task_list[new_index].priority
            priority = moved_task.priority
            if higher_priority != lower_priority:
                if (higher_priority - lower_priority) > 1:
                    priority = int(0.5 + (higher_priority + lower_priority) / 2.0)
                else:
                    priority = (higher_priority + lower_priority) / 2.0
            else:
                if new_index < len(task_list) / 2:
                    higher = find(lambda task : task.priority == higher_priority, task_list)
                    index = task_list.index(higher)
                    if index == 0:
                        priority = higher_priority + 1
                        higher_priority += 2
                    else:
                        next_priority = task_list[index - 1].priority
                        if (next_priority - higher_priority) > 2:
                            priority = higher_priority + 1
                            higher_priority += 2
                        else:
                            higher_priority += (next_priority - higher_priority) * 2.0 / 3.0
                            priority = higher_priority - (next_priority - higher_priority) / 3.0
                    for index in range(index, new_index + 1):
                        task_list[index].priority = higher_priority
                else:
                    lower = find(lambda task : task.priority == lower_priority, reversed(task_list))
                    index = task_list.index(lower)
                    if index == len(task_list) - 1:
                        priority = higher_priority - 1
                        lower_priority -= 2
                    else:
                        next_priority = task_list[index + 1].priority
                        if (lower_priority - next_priority) > 2:
                            priority = lower_priority - 1
                            lower_priority -= 2
                        else:
                            lower_priority -= (lower_priority - next_priority) * 2.0 / 3.0
                            priority = lower_priority + (lower_priority - next_priority) / 3.0
                    for index in range(new_index, index + 1):
                        task_list[index].priority = lower_priority

        moved_task.priority = priority
        self.emit('changed', 0)

    def close_cb(self, item):
        """Remove the list from the display."""
        config['todo']['show'] = False
        self.destroy()

    def key_release_cb(self, widget, event):
        """Performs the function associated with the released key."""
        if event.keyval == gtk.keysyms.Escape:
            if event.state & gtk.gdk.SHIFT_MASK:
                self.destroy()
            else:
                self.hide_cb(None)

        elif event.keyval == gtk.keysyms.Page_Down:
            self.opacity -= 10
            self.opacity = max(self.opacity, 0)
            self.set_opacity(self.opacity / 100.0)
            config['todo']['opacitiy'] = int(self.opacity + 0.5)

        elif event.keyval == gtk.keysyms.Page_Up:
            self.opacity += 10
            self.opacity = min(self.opacity, 100)
            self.set_opacity(self.opacity / 100.0)
            config['todo']['opacitiy'] = int(self.opacity + 0.5)

        elif event.keyval == gtk.keysyms.Right and event.state & gtk.gdk.MOD1_MASK:
            change = {1:2, 3:4, 8:2, 9:4}
            if self.requested_position in change:
                self.requested_position = change[self.requested_position]
            else:
                self.requested_position = 7
            self.update()
            config['todo']['position'] = self.requested_position
        elif event.keyval == gtk.keysyms.Left and event.state & gtk.gdk.MOD1_MASK:
            change = {2:1, 4:3, 8:1, 9:3}
            if self.requested_position in change:
                self.requested_position = change[self.requested_position]
            else:
                self.requested_position = 6
            self.update()
            config['todo']['position'] = self.requested_position
        elif event.keyval == gtk.keysyms.Up and event.state & gtk.gdk.MOD1_MASK:
            change = {3:1, 4:2, 6:1, 7:2}
            if self.requested_position in change:
                self.requested_position = change[self.requested_position]
            else:
                self.requested_position = 8
            self.update()
            config['todo']['position'] = self.requested_position
        elif event.keyval == gtk.keysyms.Down and event.state & gtk.gdk.MOD1_MASK:
            change = {1:3, 2:4, 6:3, 7:4}
            if self.requested_position in change:
                self.requested_position = change[self.requested_position]
            else:
                self.requested_position = 9
            self.update()
            config['todo']['position'] = self.requested_position

    def redisplay_cb(self):
        """Unhide the window and update it's contents."""
        self.allow_updates = True
        self.show_all()
        self.update()
        return False

    def sort_list(self, task):
        """Return a tuple containing the elements used to sort the task list."""
        result = []
        for key in self.key_list:
            if key == 'name' and self.show_full_name:
                result.append(task.full_name())
            else:
                value = task.__getattribute__(key)
                if not value and 'date' in key:
                    value = datetime.datetime.min
                result.append(value)
        result.append(task)
        return result

    def update_todo_list(self):
        """Regenerate the list of task used to form the todo list."""
        if config['todo'].get('sort_key') != 'recent':
            if self.key_list[0] != config['todo'].get('sort_key'):
                self.key_list.insert(0, config['todo'].get('sort_key'))
                self.key_list = self.key_list[:5]

        if config['todo'].get('sort_key') == 'recent' and self.session.task_manager.recent_tasks:
            task_list = self.session.task_manager.recent_tasks
        else:
            task_list = [self.sort_list(task) for task
                         in expand_list(self.session.task_manager.tasks)
                         if task.is_runable()]
            if self.key_list[0] == 'priority':
                task_list.sort(reverse=True)
            else:
                task_list.sort()
            task_list = [task[-1] for task in task_list]
        if not config['todo'].get('scheduled'):
            task_list = [task for task in task_list if task.state != 'Scheduled']
        if self.session.current_activity and self.current_first:
            while self.session.current_activity.task in task_list:
                task_list.remove(self.session.current_activity.task)
            task_list.insert(0, self.session.current_activity.task)

        if self.list_size:
            task_list = task_list[:self.list_size]

        for task in task_list:
            text = task.full_name() if self.show_full_name else task.name

            if self.session.current_activity and self.session.current_activity.task == task:
                text = '<b>%s</b>' % (text)

            self.list.append((weakref.ref(task), text))

    def update(self, update_menu=True):
        """Update the display."""
        if self.allow_update:
            self.list.handler_block(self.change_handler)
            self.list.clear()
            self.resize(10, 10)
            self.update_todo_list()
            width, height = self.get_size()
            x, y = self.get_position()
            screen_width = gtk.gdk.screen_width()
            screen_height = gtk.gdk.screen_height()

            menu = self.menu.get_children()[6].get_submenu().get_children()

            if self.requested_position == 0:
                if update_menu:
                    [item.set_inconsistent(True) for item in menu if item.get_active()]
            elif self.requested_position == 1:
                self.move(0, 0)
                if update_menu:
                    menu[0].set_active(True)
            elif self.requested_position == 2:
                self.move(screen_width - width, 0)
                if update_menu:
                    menu[1].set_active(True)
            elif self.requested_position == 3:
                self.move(0, screen_height - height)
                if update_menu:
                    menu[2].set_active(True)
            elif self.requested_position == 4:
                self.move(screen_width - width, screen_height - height)
                if update_menu:
                    menu[3].set_active(True)
            elif self.requested_position == 5:
                self.move((screen_width - width) / 2, (screen_height - height) / 2)
                if update_menu:
                    menu[4].set_active(True)
            elif self.requested_position == 6:
                self.move(0, y)
                if update_menu:
                    [item.set_inconsistent(True) for item in menu if item.get_active()]
            elif self.requested_position == 7:
                self.move(screen_width - width, y)
                if update_menu:
                    [item.set_inconsistent(True) for item in menu if item.get_active()]
            elif self.requested_position == 8:
                self.move(x, 0)
                if update_menu:
                    [item.set_inconsistent(True) for item in menu if item.get_active()]
            elif self.requested_position == 9:
                self.move(x, screen_height - height)
                if update_menu:
                    [item.set_inconsistent(True) for item in menu if item.get_active()]
            self.list.handler_unblock(self.change_handler)

    def redraw(self):
        """Regenerate the window."""
        self.update()

class Task_Tree(gtk.ScrolledWindow):
    """Task editor GUI class"""

    def __init__ (self, manager, callback):
        """Initialise the task application GUI."""
        gtk.ScrolledWindow.__init__(self)
        self.set_policy(gtk.POLICY_AUTOMATIC, gtk.POLICY_AUTOMATIC)

        self.note_image = gtk.gdk.pixbuf_new_from_file(config.get('image_root') + '16x16/mimetypes/text-plain.png')

        columns = [('Tasks Name', str, True),
                   ('Booking Number', str, True),
                   ('% Complete', str, False),
                   ('Start Date', str, True),
                   ('End Date', str, True),
                   ('Allocated Hours', str, True),
                   ('Sub-total Hours', str, False),
                   ('Estimated Hours', str, True),
                   ('Priority', str, True),
                   ('State', gobject.TYPE_STRING, True),
                   ('Notes', gtk.gdk.Pixbuf, False)]

        self.allow_updates = True

        column_type_list = [object]
        column_type_list += [item[1] for item in columns]
        self.store = gtk.TreeStore(*column_type_list)
        self.view = gtk.TreeView(self.store)
        self.view.get_selection().set_mode(gtk.SELECTION_BROWSE)

        self.state_list = gtk.ListStore(gobject.TYPE_STRING)
        for state in Task.states:
            self.state_list.append([state])

        count = 0
        for heading, data_type, editable in columns:
            if data_type == str:
                text_renderer = gtk.CellRendererText()
                text_renderer.set_property('editable', editable)
                text_renderer.connect('edited', self.edited_cb, count)
                text_renderer.connect('editing-started', self.editing_cb)
                text_renderer.connect('editing-canceled', self.editing_canceled_cb)
                tree_column = gtk.TreeViewColumn(heading, text_renderer, text = count + 1)
            elif data_type == gobject.TYPE_STRING:
                combo_renderer = gtk.CellRendererCombo()
                combo_renderer.set_property('text-column', 0)
                combo_renderer.set_property('editable', editable)
                combo_renderer.connect('edited', self.state_edited_cb)
                combo_renderer.connect('editing-started', self.editing_cb)
                combo_renderer.connect('editing-canceled', self.editing_canceled_cb)
                combo_renderer.set_property('has-entry', False)
                combo_renderer.set_property('model', self.state_list)
                tree_column = gtk.TreeViewColumn(heading, combo_renderer, text = count + 1)
            elif data_type == bool:
                toggle_renderer = gtk.CellRendererToggle()
                toggle_renderer.set_property('activatable', editable)
                toggle_renderer.connect('toggled', self.toggled_cb, count)
                tree_column = gtk.TreeViewColumn(None, toggle_renderer, active = count + 1)
            elif data_type == gtk.gdk.Pixbuf:
                image_renderer = gtk.CellRendererPixbuf()
                tree_column = gtk.TreeViewColumn(heading, image_renderer)
                tree_column.set_cell_data_func(image_renderer, self.icon_function)
            count += 1
            tree_column.set_sizing(gtk.TREE_VIEW_COLUMN_AUTOSIZE)
            self.view.append_column(tree_column)
        self.add(self.view)

        self.task_manager = manager
        self.callback = callback
        self.show_inactive = True

        task_list = manager.tasks
        if config.get('sort_tasks'):
            task_list.sort()
        if config.get('submenus_on_top'):
            sub_menus = [child for child in task_list if child.children]
            elements = [child for child in task_list if not child.children]
            task_list = sub_menus + elements

        for task in task_list:
            self.add_task_to_store(task)

        self.view.set_reorderable(True)
        self.view.expand_all()

        self.connect('delete-event', self.delete_event_cb)
        self.connect('popup-menu', self.context_menu_cb)
        self.view.connect('button-press-event', self.button_cb)
        self.view.connect('row_activated', self.row_activated_cb)
        self.key_handler = self.view.connect('key-release-event', self.key_release_cb)
        self.change_handler = self.store.connect_after('row-changed', self.row_changed_cb)

        self.show_all()

    def icon_function(self, column, cell, treestore, iter):
        """Display the object name"""
        task = treestore.get_value(iter, 0)()
        if task.notes:
            cell.set_property('pixbuf', self.note_image)
        else:
            cell.set_property('pixbuf', None)

    def row_activated_cb(self, view, path, view_column):
        """Display the task dialog for the task in the activated row."""
        iter = self.store.get_iter(path)
        task = self.store.get_value(iter, 0)()
        Task_Dialog(self.task_manager, None, task)

    def editing_cb(self, cell_renderer, editable, path):
        """Disable updates while editing cells."""
        self.view.handler_block(self.key_handler)
        self.allow_updates = False

    def editing_canceled_cb(self, cell_renderer):
        """Re-enables updates."""
        self.view.handler_unblock(self.key_handler)
        self.allow_updates = True

    def edited_cb(self, cell_renderer, path, new_text, column):
        """Update the task data associated with the edited field."""
        self.view.handler_unblock(self.key_handler)
        self.allow_updates = True
        iter = self.store.get_iter(path)
        task = self.store.get_value(iter, 0)()
        self.store.set_value(iter, column + 1, new_text)
        if not task:
            return
        function = [self.set_name,
                    self.set_booking_number,
                    None,
                    self.set_start_date,
                    self.set_end_date,
                    self.set_allocated_time,
                    None,
                    self.set_estimated_time,
                    self.set_priority][column]
        update = function(task, new_text)
        if update:
            self.store.set_value(iter, column + 1, update)
        self.task_manager.write_task_list()

    def state_edited_cb(self, renderer, path, new_text):
        """Responds to the task state being edited."""
        iter = self.store.get_iter(path)
        task = self.store.get_value(iter, 0)()
        task.state = new_text
        self.store.set(iter, 10, task.state)

    def toggled_cb(self, cell_renderer, path, column):
        """Update the active state of the task associated with the field."""
        iter = self.store.get_iter(path)
        task = self.store.get_value(iter, 0)()
        new_state = not self.store.get_value(iter, column + 1)
        self.store.set_value(iter, column + 1, new_state)
        if not task:
            return
        task.state = 'active' if new_state else 'inactive'
        if self.callback:
            self.callback()

    def row_changed_cb(self, model, path, iter):
        """Moves the task as per user drag and drop action."""
        task = model.get_value(iter, 0)()
        parent_iter = model.iter_parent(iter)
        if parent_iter:
            parent = model.get_value(parent_iter, 0)()
        else:
            parent = None

        if task.parent != parent:
            self.task_manager.reparent(task, parent)

        new_list = []
        child_iter = model.iter_children(parent_iter)
        while child_iter:
            this_task = model.get_value(child_iter, 0)()
            if this_task != task or model.get_path(child_iter) == path:
                new_list.append(this_task)
            child_iter = model.iter_next(child_iter)
        self.task_manager.update_children(parent, new_list)

        self.allow_updates = False
        if self.callback:
            self.callback()
        self.allow_updates = True

    def key_release_cb(self, widget, event):
        """Performs the function associated with the released key."""
        if event.keyval not in [gtk.keysyms.Delete,
                                gtk.keysyms.Insert]:
            return False

        selection = self.view.get_selection()
        selected = selection.get_selected()[1]

        if selected and event.keyval == gtk.keysyms.Delete:
            task = self.store.get_value(selected, 0)()
            self.delete_task(task, selected)

        if selected and event.keyval == gtk.keysyms.Insert:
            path_to_task = self.store.get_path(selected)
            if event.state & gtk.gdk.SHIFT_MASK:
                self.insert_task(parent_iter = selected)
            else:
                self.insert_task(sibling_iter = selected)
            self.view.expand_to_path(path_to_task)

    def insert_task(self, parent_iter=None, sibling_iter=None):
        """Inserts a new task into the tree."""
        self.store.handler_block(self.change_handler)

        if sibling_iter:
            prior = self.store.get_value(sibling_iter, 0)()
            parent_iter = self.store.iter_parent(sibling_iter)
        else:
            prior = None

        if parent_iter:
            parent = self.store.get_value(parent_iter, 0)()
            task_names = [child.name for child in parent.children]
        else:
            parent = None
            task_names = [task.name for task in self.task_manager.tasks]

        path = 'New Task'
        count = 1
        while (path + str(count)) in task_names:
            count += 1
        path = path + str(count)
        task = self.task_manager.new_task(path, prior=prior, parent=parent)
        if task:
            self.add_task_to_store(task, parent_iter, sibling_iter)

        if self.callback:
            self.callback()

        self.store.handler_unblock(self.change_handler)

    def delete_task(self, task, iter):
        """Performs the action the via the task manager."""
        self.store.handler_block(self.change_handler)

        result = None
        if task.get_hours_subtotal() > 0.0:
            Message.show('Cannot Delete Task',
                         "Task %s cannot be deleted because it or one of it's child tasks has allocated to it.\n" % (task.name) +
                         'Only tasks with no time allocated to them can be deleted.',
                         gtk.MESSAGE_WARNING)
        else:
            Confirmation("Delete Task?",
                         "Really delete %s and children?" % (task.name),
                         self.delete_task_cb, (task, iter))

        self.store.handler_unblock(self.change_handler)

    def delete_task_cb(self, data):
        """Delete the task following confirmation."""
        task = data[0]
        task_iter = data[1]
        self.task_manager.delete_task(task)
        if self.store.remove(task_iter):
            self.view.get_selection().select_iter(task_iter)
        if self.callback:
            self.callback()

    def menu_delete_cb(self, widget):
        """Delete the task following confirmation."""
        task, task_iter = self.get_selected_task()
        if task:
            self.delete_task(task, task_iter)

    def button_cb(self, widget, event):
        """If right button display context menu, alse perform standard event processing."""
        if event.type == gtk.gdk.BUTTON_PRESS and event.button == 3:
            self.context_menu_cb(widget)

    def context_menu_cb(self, widget):
        """Displays the context menu for the task tree."""
        menu = build_menu(self.create_menu_definition())
        menu.popup(None, None, None, 1, gtk.get_current_event_time())

    def task_displayed(self, task):
        """Returns a flag indicating whether the referenced task should be displayed."""
        result = False
        if task.state != 'deleted':
            if self.show_inactive or task.has_runable_child():
                result = True
        return result

    def add_task_to_store(self, task, parent=None, sibling=None):
        """Adds the specified task and it's children to the tree store."""
        iter = None
        if self.task_displayed(task):
            if task.start_date:
                start = str(task.start_date.strftime(config.get('date_format')))
            else:
                start = 'Not set'
            if task.end_date:
                end = str(task.end_date.strftime(config.get('date_format')))
            else:
                end = 'Not set'
            allocated_time = '%02.02f' % (get_decimal_hours(task.allocated_time))
            hours_subtotal = '%02.02f' % (task.get_hours_subtotal())
            estimated_time = '%02.02f' % (get_decimal_hours(task.estimated_time))
            progress = '%0.2f' % (task.get_progress())
            iter = self.store.insert_before(parent,
                                            sibling,
                                            (weakref.ref(task),
                                             task.name,
                                             task.booking_number,
                                             progress,
                                             start,
                                             end,
                                             allocated_time,
                                             hours_subtotal,
                                             estimated_time,
                                             str(task.priority),
                                             task.state,
                                             None))

            task_list = task.children
            if config.get('sort_tasks'):
                task_list.sort()
            if config.get('submenus_on_top'):
                sub_menus = [child for child in task_list if child.children]
                elements = [child for child in task_list if not child.children]
                task_list = sub_menus + elements
            for child in task_list:
                self.add_task_to_store(child, iter)
        return iter

    def check_tasks(self, task_list, parent_iter):
        """Checks the tree model against a task list updating the model as necessary."""
        start_iter = self.store.iter_children(parent_iter)
        list_iter = start_iter
        if len(task_list) > 0:
            mapping = {}
            while list_iter:
                task = self.store.get_value(list_iter, 0)()
                mapping[task] = list_iter
                list_iter = self.store.iter_next(list_iter)

            previous = None
            list_iter = start_iter
            for task in task_list:
                stored_task = None
                if list_iter:
                    stored_task = self.store.get_value(list_iter, 0)()
                if task != stored_task:
                    if task in mapping:
                        self.store.move_after(mapping[task], previous)
                        if previous != None:
                            list_iter = self.store.iter_next(previous)
                        else:
                            list_iter = mapping[task]
                    else:
                        list_iter = self.add_task_to_store(task, parent = parent_iter, sibling = list_iter)
                tasks = [task for task in task.children if self.task_displayed(task)]
                self.check_tasks(tasks, list_iter)
                previous = list_iter
                if previous:
                    list_iter = self.store.iter_next(previous)
            if list_iter:
                while self.store.iter_is_valid(list_iter):
                    self.store.remove(list_iter)
        elif start_iter:
            iter = start_iter
            while self.store.iter_is_valid(iter):
                self.store.remove(iter)

    def check_tree(self):
        """Checks the tree model against the task hierarchy updating the model as necessary."""
        if not self.allow_updates:
            return
        self.store.handler_block(self.change_handler)

        task_list = [task for task in self.task_manager.tasks if self.task_displayed(task)]
        iter = self.store.get_iter_first()

        self.check_tasks(task_list, None)
        self.update()
        self.store.handler_unblock(self.change_handler)

    def update(self):
        """Update values displayed in the widget."""
        if not self.allow_updates:
            return
        self.store.handler_block(self.change_handler)
        iter_stack = [self.store.get_iter_first()]
        iter = iter_stack[-1]
        while iter:
            task = self.store.get_value(iter, 0)()
            if task.start_date:
                start = str(task.start_date.strftime(config.get('date_format')))
            else:
                start = 'Not set'
            if task.end_date:
                end = str(task.end_date.strftime(config.get('date_format')))
            else:
                end = 'Not set'
            allocated_time = '%02.02f' % (get_decimal_hours(task.allocated_time))
            hours_subtotal = '%02.02f' % (task.get_hours_subtotal())
            estimated_time = '%02.02f' % (get_decimal_hours(task.estimated_time))
            progress = '%0.2f' % (task.get_progress())
            self.store.set(iter, 1, task.name,
                                 2, task.booking_number,
                                 3, progress,
                                 4, start,
                                 5, end,
                                 6, allocated_time,
                                 7, hours_subtotal,
                                 8, estimated_time,
                                 9, task.priority,
                                 10, task.state)
            child = self.store.iter_children(iter)
            if child:
                iter_stack.append(iter)
                iter = child
            else:
                iter = self.store.iter_next(iter)
                if not iter and iter_stack:
                    iter = iter_stack.pop()
                    iter = self.store.iter_next(iter)
        self.store.handler_unblock(self.change_handler)

    def redraw(self):
        """Redraw the widget."""
        if not self.allow_updates:
            return
        self.check_tree()

    def create_menu_definition(self):
        """Returns the definition used to create the context menu."""
        menu = []
        selection = self.view.get_selection()
        selected = selection.get_selected()[1]
        if selected:
            menu.extend([('Edit Task', self.edit_task_cb),
                         ('---', '---'),
                         ('Add Sub-Task', self.add_subtask_cb, selected),
                         ('Add Task', self.add_task_cb, selected),
                         ('Delete Task', self.menu_delete_cb),
                         ('---', '---')])
        menu.extend([('Expand All', self.expand_all_cb),
                    ('Collapse All', self.collapse_all_cb)])

        if self.show_inactive:
            menu.append(('Hide Inactive Tasks', self.show_inactive_cb, False))
        else:
            menu.append(('Show Inactive Tasks', self.show_inactive_cb, True))

        return menu

    def get_selected_task(self):
        """Returns the task associated with the selected row, or None."""
        task = None
        selection = self.view.get_selection()
        model, iter = selection.get_selected()
        if iter:
            task = model.get_value(iter, 0)()
        return task, iter

    def expand_all_cb(self, item):
        """Expands all task in the tree to show all children and sub children etc."""
        self.view.expand_all()

    def collapse_all_cb(self, item):
        """Collapses all task in the tree to show only top level tasks."""
        self.view.collapse_all()

    def show_inactive_cb(self, item, value):
        """Hides / Shows inactive tasks based on the setting of value."""
        self.show_inactive = value
        self.redraw()

    def edit_task_cb(self, item):
        """Display task editing GUI."""
        task = self.get_selected_task()[0]
        if task:
            Task_Dialog(self.task_manager, None, task)

    def add_subtask_cb(self, item, parent_iter):
        """Add task as a subtask of parent."""
        path_to_task = self.store.get_path(parent_iter)
        self.insert_task(parent_iter = parent_iter)
        self.view.expand_to_path(path_to_task)

    def add_task_cb(self, item, sibling_iter):
        """Add task as a sibling."""
        path_to_task = self.store.get_path(parent_iter)
        self.insert_task(sibling_iter = sibling_iter)
        self.view.expand_to_path(path_to_task)

    def set_name(self, task, name):
        task.name = name

    def set_booking_number(self, task, booking_number):
        task.booking_number = booking_number

    def set_start_date(self, task, date):
        result = "Not set"
        if date:
            task.start_date = string_to_date(date)
            if task.start_date:
                result = task.start_date.strftime(config.get('date_format'))
        return result

    def set_end_date(self, task, date):
        result = "Not set"
        if date:
            task.end_date = string_to_date(date)
            if task.end_date:
                result = task.end_date.strftime(config.get('date_format'))
        return result

    def set_allocated_time(self, task, time):
        task.allocated_time = datetime.timedelta(hours = float(time))
        return '%02.2f' % get_decimal_hours(task.allocated_time)

    def set_estimated_time(self, task, time):
        task.estimated_time = datetime.timedelta(hours = float(time))
        return '%02.2f' % get_decimal_hours(task.estimated_time)

    def set_priority(self, task, priority):
        task.priority = int(priority)
        return '%d' % task.priority

    def delete_event_cb(self, widget, event):
        # Change False to True and the main window will not be destroyed
        self.task_manager.write_task_list()

class Bookings(gtk.Frame):
    """Widget containing the booking totals for a week."""

    def __init__(self, session):
        """Initialise the booking widget from the manager data."""
        gtk.Frame.__init__(self)
        self.bookings = None
        self.session = session
        self.set_shadow_type(gtk.SHADOW_NONE)
        self.box = gtk.VBox()
        button_box = gtk.HBox()
        show_all = gtk.CheckButton('Show Unbilled Tasks')
        show_all.set_active(config.get('show_unbilled'))
        show_all.connect('toggled', self.show_all_toggled_cb)
        include_future = gtk.CheckButton('Show Future Tasks')
        include_future.set_active(config.get('include_future'))
        include_future.connect('toggled', self.include_future_toggled_cb)
        button_box.pack_start(show_all, False, False)
        button_box.pack_start(include_future, False, False)
        self.box.pack_start(button_box, False, False)
        self.draw()
        self.add(self.box)

    def show_all_toggled_cb(self, button):
        """Processes the state of the show all button."""
        config['show_unbilled'] = button.get_active()
        self.redraw()

    def include_future_toggled_cb(self, button):
        """Processes the state of the show all button."""
        config['include_future'] = button.get_active()
        self.redraw()

    def update(self):
        """Updates the booking number values."""
        self.redraw()

    def redraw(self):
        """Redraw the entire booking number display."""
        if self.bookings in self.box.get_children():
            self.box.remove(self.bookings)
        self.draw()

    def draw(self):
        """Displays the booking details."""
        allocated_tasks = set()
        booking_numbers = set()

        now = datetime.datetime.now()
        for activity in self.session.activities:
            if activity.allocation and (activity.start_event.time < now
                                        or config.get('include_future')):
                booking_number = activity.task.get_booking_number()
                if booking_number[0] != '#' or config.get('show_unbilled'):
                    allocated_tasks.add(activity)
                    booking_numbers.add(booking_number)

        num_lines = len(booking_numbers)
        
        if num_lines:
            week_start = self.session.week_start

            self.bookings = gtk.Table(9, num_lines + 2, False)
            self.box.pack_end(self.bookings)

            row = 1
            totals = {}
            tips = {}
            rows = {}
            for day in range(0, 7):
                title = gtk.Label((week_start + datetime.timedelta(day)).strftime('%A'))
                title.set_width_chars(9)
                self.bookings.attach(title, day + 1, day + 2, 0, 1, gtk.FILL|gtk.EXPAND, gtk.FILL, 0, 0)
            title = gtk.Label("Total")
            title.set_width_chars(9)
            self.bookings.attach(title, 8, 9, 0, 1, gtk.FILL|gtk.EXPAND, gtk.FILL, 0, 0)
            for booking_number in booking_numbers:
                if not booking_number.startswith('#'):
                    label = gtk.Label(booking_number)
                else:
                    if booking_number.startswith('#/'):
                        text = booking_number[2:]
                    else:
                        text = booking_number[1:]
                    label = gtk.Label()
                    label.set_markup("<i>%s</i>" % (text))
                self.bookings.attach(label, 0, 1, row, row + 1, 0, 0, 0, 0)
                totals[booking_number] = [0.0] * 7
                tips[booking_number] = [''] * 7
                if booking_number not in rows:
                    rows[booking_number] = row
                row += 1
            self.bookings.attach(gtk.Label("Total"), 0, 1, row, row + 1, 0, 0, 0, 0)

            if config.get('include_future'):
                events = [event for event in self.session.events]
            else:
                events = [event for event in self.session.events if event.time < now]

            if (self.session.week_start < now and
                self.session.week_start + datetime.timedelta(7) > now):
                events.append(Now(self.session))
            events.sort(key=operator.attrgetter('time'))

            accumulated = set()
            unaccumulated = set()
            previous_event = None
            for event in events:
                # Ignore events that do not change activity states except 'Now'
                # event
                if event.get_id() and not event.prior_activity and not \
                    event.subsequent_activity:
                    continue

                count = len(accumulated)
                day = event.time.weekday()
                ratio = sum([activity.allocation for activity in accumulated])
                for activity in accumulated:
                    index = activity.task.get_booking_number()
                    hours = get_decimal_hours(event.time - previous_event.time)
                    hours *= activity.allocation / ratio
                    totals[index][day] += hours
                    if config.get('booking_tips_full'):
                        name = activity.task.full_name()
                    else:
                        name = activity.task.name
                    if ratio > 1.0:
                        tips[index][day] += '%s: %02.02f (%02.02f / %02.01f)\n' % (name,
                                                                                   hours,
                                                                                   hours * ratio,
                                                                                   ratio)
                    else:
                        tips[index][day] += '%s: %02.02f\n' % (name, hours)
                for activity in unaccumulated:
                    hours = get_decimal_hours(event.time - previous_event.time)
                    index = activity.task.get_booking_number()
                    totals[index][day] += hours
                    if config.get('booking_tips_full'):
                        name = activity.task.full_name()
                    else:
                        name = activity.task.name
                    tips[index][day] += '%s: %02.02f\n' % (name, hours)

                for prior in event.prior_activity:
                    if prior.allocation == 0:
                        continue
                    if not previous_event:
                        Message.show("Inconsistent State",
                                     "Event %d has prior activity with no prior event!" % (event.get_id()),
                                     gtk.MESSAGE_WARNING)

                    # Handle case of current activity having an end event after
                    # now by ignoring 'Now' event when future activity is
                    # included.
                    #
                    if config.get('include_future') and event.get_id() == 0 and prior.end_event != event:
                        continue  

                    if prior in accumulated:                        
                        accumulated.remove(prior)
                    elif prior in unaccumulated:
                        unaccumulated.remove(prior)
                    elif prior not in allocated_tasks:
                        pass
                    elif config.get('show_unbilled') or not prior.task.get_booking_number().startswith('#'):
                        Message.show("Inconsistent State",
                                     ("Error processing event %d\nActivity %d " +
                                      "is not in accumulated or unaccumulated list!") % (event.get_id(),
                                                                                         prior.get_id()),
                                     gtk.MESSAGE_WARNING)

                if config.get('include_future') or event.time <= now:
                    for subsequent in event.subsequent_activity:
                        if subsequent.allocation == 0:
                            continue
                        if not subsequent.task.get_booking_number().startswith('#'):
                            accumulated.add(subsequent)
                        elif config.get('show_unbilled'):
                            unaccumulated.add(subsequent)
                previous_event = event

            for day in range(0, 7):
                day_total = 0
                for key in totals.keys():
                    total = totals[key][day]
                    # Skip elements starting with '#'
                    if str(key)[0] != "#":
                        day_total += totals[key][day]
                        title = gtk.Label("%0.2f" % (total))
                    else:
                        title = gtk.Label()
                        title.set_markup("<i>%0.2f</i>" % (total))
                    tip = tips[key][day].strip()
                    if tip:
                        box = gtk.EventBox()
                        box.add(title)
                        box.set_tooltip_text(tip)
                        box.add_events(gtk.gdk.BUTTON_PRESS_MASK)
                        box.connect('button-press-event', self.key_cb)
                        title = box
                    row = rows[key]
                    frame = gtk.Frame()
                    frame.add(title)
                    self.bookings.attach(frame,
                                         day + 1,
                                         day + 2,
                                         row, row + 1,
                                         gtk.FILL|gtk.EXPAND,
                                         gtk.FILL)
                title = gtk.Label("%0.2f" % (day_total))
                frame = gtk.Frame()
                frame.add(title)
                self.bookings.attach(frame,
                                     day + 1,
                                     day + 2,
                                     num_lines + 1,
                                     num_lines + 2,
                                     gtk.FILL|gtk.EXPAND,
                                     gtk.FILL)
                
            grand_total = 0.0
            for key in totals.keys():
                total = reduce(operator.add, totals[key])
                # Skip elements starting with '#'
                if str(key)[0] != "#":
                    grand_total += total
                text = gtk.Label("%0.2f" % (total))
                frame = gtk.Frame()
                frame.add(text)
                row = rows[key]
                self.bookings.attach(frame,
                                     8,
                                     9,
                                     row,
                                     row + 1,
                                     gtk.FILL|gtk.EXPAND,
                                     gtk.FILL)

            text = gtk.Label("%0.2f" % (grand_total))
            frame = gtk.Frame()
            frame.add(text)
            self.bookings.attach(frame,
                                 8,
                                 9,
                                 num_lines + 1,
                                 num_lines + 2,
                                 gtk.FILL|gtk.EXPAND,
                                 gtk.FILL)
            self.box.show_all()

    def key_cb(self, widget, event):
        """Processes a key press on the booking sheet."""
        text = widget.get_tooltip_text()
        gtk.Clipboard().set_text(text)

class Timesheet(gtk.Table):
    """Widget containing a timesheet."""

    def __init__(self, session, callback):
        """Initialise the display of a timesheet page."""
        gtk.Table.__init__(self, 3, 3, False)
        self.session = session
        self.callback = callback
        self.row_height = 50

        self.draw()

    def update(self):
        """Update the timesheet display."""
        for day in range(0, 7):
            self.task_timeline[day].queue_draw()

    def redraw(self):
        """Update the timesheet display."""
        week_start = self.session.week_start
        self.scale_line.queue_draw()

        for day in range(0, 7):
            self.task_timeline[day].session = self.session

            day_start = (week_start + datetime.timedelta(day))
            day_end = (week_start + datetime.timedelta(day + 1))
            self.task_timeline[day].date = day_start.date()

            self.day_label[day].set_text(day_start.strftime("%A\n%d/%m/%Y"))
            activities = []
            for activity in self.session.activities:
                if not activity.task:
                    continue
                start = activity.start_event.time if activity.start_event else datetime.datetime.now()
                end = activity.end_event.time if activity.end_event else datetime.datetime.now()
                if start >= day_start and end <= day_end:
                    activities.append(activity)
            self.task_timeline[day].update_activity_list(activities)
            self.task_timeline[day].queue_draw()

    def draw(self):
        """Initialise the display of a timesheet page."""
        week_start = self.session.week_start.replace(hour = 0, minute = 0, second = 0, microsecond = 0)

        zoom_buttons = gtk.VButtonBox()
        zoom_in = gtk.Button(stock = gtk.STOCK_ZOOM_IN)
        zoom_in.connect('clicked', self.zoom_in_cb)
        zoom_out = gtk.Button(stock = gtk.STOCK_ZOOM_OUT)
        zoom_out.connect('clicked', self.zoom_out_cb)
        zoom_buttons.pack_start(zoom_in)
        zoom_buttons.pack_start(zoom_out)
        self.attach(zoom_buttons, 0, 1, 0, 1, 0, 0, 0, 0)

        vertical_scrollbar = gtk.VScrollbar()
        self.attach(vertical_scrollbar,
                    2,
                    3,
                    1,
                    8,
                    gtk.FILL,
                    gtk.FILL | gtk.SHRINK)
        horizontal_scrollbar = gtk.HScrollbar()
        self.attach(horizontal_scrollbar,
                    1,
                    2,
                    8,
                    9,
                    gtk.FILL | gtk.SHRINK,
                    gtk.FILL)
        horizontal_scrollbar.connect('value-changed',
                                     self.update_position_cb)
        now = datetime.datetime.now()
        work_day_start = now.replace(hour = int(config.get('day_start')[:2]),
                                     minute = int(config.get('day_start')[3:]),
                                     second = 0,
                                     microsecond = 0)
        work_day_end = now.replace(hour = int(config.get('day_end')[:2]),
                                   minute = int(config.get('day_end')[3:]),
                                   second = 0,
                                   microsecond = 0)
        self.scale_line = Timeline(now.date())
        self.scale_line.set_hadjustment(horizontal_scrollbar.get_adjustment())
        self.scale_line.editable = False
        start = Event(self.session)
        start.time = work_day_start.replace(minute = 0,
                                            second = 0,
                                            microsecond = 0)
        for hour in range(work_day_start.hour, work_day_end.hour):
            end = Event(self.session)
            end.time = start.time + datetime.timedelta(hours = 1)
            text = "%02d:00-%02d:00" % (hour, hour + 1)
            self.scale_line.new_activity(None, start, end, text)
            start = end
        self.attach(self.scale_line, 1, 2, 0, 1, gtk.FILL, gtk.FILL)

        timeline_table = gtk.Table(1, 7)
        self.scrolled_timelines = gtk.Layout(vadjustment = vertical_scrollbar.get_adjustment())
        self.scrolled_timelines.set_size(100, int((self.row_height * 7) + 0.5))
        self.scrolled_timelines.put(timeline_table, 0, 0)
        self.scrolled_timelines.connect("expose-event", self.expose_cb)

        day_table = gtk.Table(1, 7)
        self.scrolled_days = gtk.Layout(vadjustment = vertical_scrollbar.get_adjustment())
        self.scrolled_days.set_size(100, int((self.row_height * 7) + 0.5))
        self.scrolled_days.put(day_table, 0, 0)
        self.scrolled_days.connect("expose-event", self.expose_cb)

        self.attach(self.scrolled_days, 0, 1, 1, 2, gtk.FILL, gtk.FILL | gtk.EXPAND)
        self.attach(self.scrolled_timelines, 1, 2, 1, 2, gtk.FILL | gtk.EXPAND, gtk.FILL | gtk.EXPAND)
        self.task_timeline = [None] * 7
        self.day_label = [None] * 7

        for day in range(0, 7):
            day_start = (week_start + datetime.timedelta(day))
            day_end = (week_start + datetime.timedelta(day + 1))
            frame = gtk.Frame()
            frame.set_size_request(-1, int(self.row_height + 0.5))
            self.day_label[day] = gtk.Label(day_start.strftime("%A\n%d/%m/%Y"))
            self.day_label[day].set_justify(gtk.JUSTIFY_CENTER)

            frame.add(self.day_label[day])
            day_table.attach(frame, 0, 1, day, day + 1)
            self.task_timeline[day] = Timeline(day_start.date(), self.session, self.callback, self.scale_line)
            for activity in self.session.activities:
                if not activity.task:
                    continue
                start = activity.start_event.time if activity.start_event else datetime.datetime.now()
                end = activity.end_event.time if activity.end_event else datetime.datetime.now()
                if start > day_start and end < day_end:
                    self.task_timeline[day].new_activity(activity, activity.start_event, activity.end_event, activity.task.name)
            timeline_table.attach(self.task_timeline[day], 0, 1, day, day + 1, gtk.FILL | gtk.EXPAND, gtk.FILL | gtk.EXPAND, 0, 0)

    def adjustment_changed_cb(self, adjustment):
        """Alters the adjustment value to keep the display centered."""
        ratio = adjustment.upper / self.old_upper
        value = (adjustment.value + adjustment.page_size / 2.0) * ratio
        value -= adjustment.page_size / 2.0
        adjustment.set_value(value)
        self.old_upper = adjustment.upper

    def update_position(self, value, upper, width):
        """Alters the adjustment value to keep the display centered."""
        adjustment = self.scale_line.get_hadjustment()
        ratio = width / upper
        value = (value + adjustment.page_size / 2.0) * ratio
        value -= adjustment.page_size / 2.0
        adjustment.set_value(value)

    def update_position_cb(self, scrollbar):
        """Redraw the timelines when the position has changed."""
        self.redraw()

    def expose_cb(self, widget, event = None):
        """Redraw the vertical scrolling elements."""
        if event.window.get_geometry()[3] > self.row_height * 7:
            self.row_height = event.window.get_geometry()[3] / 7.0
            self.scrolled_timelines.set_size(100, int((self.row_height * 7) + 0.5))
            self.scrolled_days.set_size(100, int((self.row_height * 7) + 0.5))

        widget.get_vadjustment().upper = self.row_height * 7
        widget.allocation.height = self.row_height * 7
        for child in widget.get_children():
            child.set_size_request(widget.allocation.width,
                                   int((self.row_height * 7) + 0.5))

    def update_width(self, width):
        """Updates the width of all the timeline widgets."""
        value = self.scale_line.get_hadjustment().value
        upper = self.scale_line.get_hadjustment().upper
        self.update_position(value, upper, width)
        self.scale_line.set_size(int(width + 0.5), int(self.row_height + 0.5))
        for day in range(0, 7):
            self.task_timeline[day].set_size(int(width + 0.5),
                                             int(self.row_height + 0.5))

    def zoom_in_cb(self, widget):
        """Magnifies the timeline displays."""
        self.scale_line.zoom *= 2.0
        if self.scale_line.zoom > 64.0:
            self.scale_line.zoom = 64.0
        else:
            width = self.scale_line.get_size()[0] * 2.0
            self.update_width(width)
            self.scale_line.ratio *= 2.0
        self.redraw()

    def zoom_out_cb(self, widget):
        """Un-magnifies the timeline displays."""
        self.scale_line.zoom /= 2.0
        if self.scale_line.zoom < 0.125:
            self.scale_line.zoom = 0.125
        else:
            width = self.scale_line.get_size()[0] / 2.0
            self.update_width(width)
            self.scale_line.ratio /= 2.0
        self.redraw()

class Timesheet_Editor(gtk.Window, Managed_Window):
    """Timesheet window class"""
    timeout = None
    x = 200
    y = 200
    width = 400
    height = 400

    __gsignals__ = {'changed' : (gobject.SIGNAL_RUN_LAST, gobject.TYPE_NONE, (gobject.TYPE_INT,))}

    def __init__(self, session, tab = None):
        """Initialise the timesheet window"""
        gtk.Window.__init__(self, gtk.WINDOW_TOPLEVEL)
        Managed_Window.__init__(self, 'timesheet')

        self.session = session

        icons = []
        for image_file in glob.glob(config.get('image_root') + "*x*/apps/preferences-system-time.png"):
            icons.append(gtk.gdk.pixbuf_new_from_file(image_file))
        self.set_icon_list(*icons)

        self.set_title("Timesheet Editor")

        self.week = datetime.datetime.now()
        self.week_select = Week_Selector()
        self.week_select.go_button.connect('clicked', self.change_week_cb)

        if not tab:
            self.timesheet = Timesheet(session, self.changed_cb)
            self.bookings = Bookings(session)
            self.task_list = Task_Tree(session.task_manager, self.changed_cb)
        box = gtk.VBox()
        self.status = gtk.Statusbar()
        box.pack_start(self.week_select, False, False)
        button_bar = gtk.HButtonBox()
        if os.name == 'nt':
            outlook_import = gtk.Button('Import from outlook')
            outlook_import.connect('clicked', session.outlook_import_cb, self.changed_cb)
            button_bar.pack_start(outlook_import)

        options = gtk.Button('Edit Options')
        options.connect('clicked', self.options_cb)
        button_bar.pack_start(options)

        update_hours = gtk.Button('Update Allocated Hours')
        update_hours.connect('clicked', self.update_hours_cb)
        button_bar.pack_start(update_hours)

        save = gtk.Button('Save')
        save.connect('clicked', self.save_cb)
        button_bar.pack_start(save)

        self.tabs = gtk.Notebook()
        self.tabs.connect('page-removed', self.page_removed_cb)
        if not tab:
            self.tabs.append_page(self.timesheet, gtk.Label('Timesheet'))
            self.tabs.append_page(self.bookings, gtk.Label('Bookings'))
            self.tabs.set_tab_detachable(self.bookings, True)
            self.tabs.append_page(self.task_list, gtk.Label('Tasks'))
            self.tabs.set_tab_detachable(self.task_list, True)
        self.tabs.set_group_id(1)

        box.pack_start(self.tabs)
        box.pack_start(button_bar, False, False)
        box.pack_start(self.status, False, False)
        self.add(box)

        self.connect("map-event", self.mapped_cb)
        self.show_all()
        save.grab_focus()

    def changed_cb(self):
        """Updates the widgets when the source data has been edited."""
        self.emit('changed', 0)

    def mapped_cb(self, window, event):
        """When first mapped select the first tabbed."""
        self.tabs.set_current_page(0)

    def change_week_cb(self, widget, data = None):
        """Changes the week."""
        date = datetime.datetime.combine(self.week_select.get_date(),
                                         datetime.time(0))
        self.session.save()
        if date in Activity_Manager.sessions:
            self.session = Activity_Manager.sessions(date)
        else:
            self.session = Activity_Manager(self.session.task_manager,
                                            self.session.task_manager.update_gui,
                                            date)
        for page_index in range(self.tabs.get_n_pages()):
            self.tabs.get_nth_page(page_index).session = self.session
        self.redraw()

    def save_cb(self, widget):
        """Saves the currently displayed session."""
        self.session.save()

    def options_cb(self, widget):
        """Display the task dialog."""
        Options_Dialog()

    def update_hours_cb(self, widget):
        """Update the task hours."""
        accumulate_hours(self.session.task_manager)

    def page_removed_cb(self, notebook, child, page_num):
        """Deletes the top level widget if the notebook has no pages."""
        if not notebook.get_n_pages():
            self.get_toplevel().destroy()

    def redraw(self):
        """Redraw the widget if displayed."""
        for page_index in range(self.tabs.get_n_pages()):
            self.tabs.get_nth_page(page_index).redraw()
        self.queue_draw()

    def update(self):
        """Update the widget if displayed."""
        for page_index in range(self.tabs.get_n_pages()):
            self.tabs.get_nth_page(page_index).update()
        return True

    def set_state(self, text):
        """Update the tool tip."""
        if self.session.current_activity:
            text = '%s'
        self.set_tooltip("Time booking system\n%s" % (text))

    def update_task_menu(self):
        return

    def new_task_cb(self, widget):
        """Creates a new task dialog."""
        Task_Dialog(self.session.task_manager, self.session)

class Task_List_Dialog(gtk.Dialog, Managed_Dialog):
    """Dialog to diplay list of tasks."""

    __gsignals__ = {'changed' : (gobject.SIGNAL_RUN_LAST, gobject.TYPE_NONE, (gobject.TYPE_INT,))}

    instance = None
    activation_action = None

    @classmethod
    def show(cls, session):
        """Display the dialog, creating a new one if it's not currently displayed."""
        if cls.instance == -1:
            return
        elif cls.instance:
            cls.instance.update()
            cls.instance.present()
        else:
            cls.instance = cls(session)

    def __init__(self, session, id):
        """Create an instance of the dialog."""
        gtk.Dialog.__init__(self, "Select Task", None, 0,
                            (gtk.STOCK_CANCEL, gtk.RESPONSE_CANCEL,
                             gtk.STOCK_OK, gtk.RESPONSE_OK))
        Managed_Dialog.__init__(self, id)

        self.connect("delete-event", self.delete_cb)
        self.connect('response', self.response_cb)
        self.session = session
        icons = []
        for image_file in glob.glob(config.get('image_root') + "*x*/actions/view-pim-tasks.png"):
            icons.append(gtk.gdk.pixbuf_new_from_file(image_file))
        self.set_icon_list(*icons)
        
        entry = gtk.Entry()
        entry.connect('changed', self.search_cb)
        self.tree = self.task_list()
        window = gtk.ScrolledWindow()
        window.set_policy(gtk.POLICY_AUTOMATIC, gtk.POLICY_AUTOMATIC)
        window.add(self.tree)
        self.search_iter = None
        self.vbox.pack_start(entry, False, False)
        self.vbox.pack_start(window)
        self.show_all()

    def update(self):
        """Do nothing - overridden by sub-classes."""
        pass

    def search_cb(self, entry):
        """Updates the list of tasks based on the search criteria."""
        results = self.session.task_manager.find_by_subject(entry.get_text())
        tasks = [key for key in results.keys() if results[key] > 0]
        if entry.get_text() and tasks:
            tasks.sort(key=Task.full_name)
            if self.search_iter:
                children = self.store.iter_children(self.search_iter)
                while children and self.store.iter_is_valid(children):
                    self.store.remove(children)
            else:
                self.search_iter = self.store.insert(None, 0, (None, 'Search Results'))
            
            iters = {None:self.search_iter}
            for task in tasks:
                ancestors = []
                parent = task
                while parent:
                    ancestors.insert(0, parent)
                    parent = parent.parent
                for ancestor in ancestors:
                    if ancestor not in iters:
                        text = ancestor.name
                        if ancestor in results:
                            text += ' ' + str(results[ancestor])
                        iters[ancestor] = self.store.append(iters[ancestor.parent],
                                                            (ancestor, text))
                    
            path = self.store.get_path(self.search_iter)
            self.tree.expand_row(path, True)
        else:
            if self.search_iter:
                self.store.remove(self.search_iter)
                self.search_iter = None

    def response_cb(self, dialog, response_id):
        """Processes the user response to the task request."""
        new_instance = self
        if response_id == gtk.RESPONSE_CANCEL:
            new_instance = None
        elif response_id == gtk.RESPONSE_OK:
            selection = self.tree.get_selection()
            selected = selection.get_selected()[1]
            if selected:
                task = self.store.get_value(selected, 0)
                string = self.store.get_value(selected, 1)
                if task and task.is_runable():
                    self.session.start_task(task)
                    self.emit('changed', 0)
                    new_instance = None
                elif string == 'New Task':
                    Task_Dialog(self.session.task_manager, self.session, start=True)
                    new_instance = None                    
        else:
            new_instance = -1

        if new_instance in [None, -1]:
            if self.__class__.instance == self:
                self.__class__.instance = new_instance
            self.destroy()

    def delete_cb(self, widget, event):
        """Deletes the dialog and removes the reference to it."""
        if self.instance == self:
            self.__class__.instance = None
        return False

    def task_list(self, tree=False, order='priority'):
        """Generate a task with top priority and recent task items."""
        self.store = gtk.TreeStore(object, str)
        list_widget = gtk.TreeView(self.store)
        list_widget.set_headers_visible(False)

        text_renderer = gtk.CellRendererText()
        tree_column = gtk.TreeViewColumn('Task', text_renderer, text=1)
        tree_column.set_sizing(gtk.TREE_VIEW_COLUMN_AUTOSIZE)
        list_widget.append_column(tree_column)

        self.store.insert(None, 0, (None, 'New Task'))
        priority_iter = self.store.insert(None, 1, (None, 'Top Priority'))
        task_list = [task for task in self.session.task_manager.tasks if task.is_runable()]
        task_list.sort(key=operator.attrgetter('priority'))
        for task in task_list[:5]:
            text = task.full_name() if config['todo'].get('full_name') else task.name
            self.store.append(priority_iter, (task, text))

        if self.session.task_manager.recent_tasks:
            recent_iter = self.store.insert(None, 2, (None, 'Most Recent'))
            for task in self.session.task_manager.recent_tasks[:5]:
                if task:
                    text = task.full_name() if config['todo'].get('full_name') else task.name
                    self.store.append(recent_iter, (task, text))

        all_iter = self.store.insert(None, 3, (None, 'All Tasks'))
        task_list = self.session.task_manager.tasks
        if config.get('sort_tasks'):
            task_list.sort()
        if config.get('submenus_on_top'):
            sub_menus = [child for child in task_list if child.children]
            elements = [child for child in task_list if not child.children]
            task_list = sub_menus + elements
        for task in task_list:
            self.add_task_to_store(task, parent=all_iter)

        list_widget.connect('row_activated', self.row_activated_cb)

        return list_widget

    def row_activated_cb(self, view, path, view_column):
        """Display the task dialog for the task in the activated row."""
        store = view.get_model()
        iter = store.get_iter(path)
        task = store.get_value(store.get_iter(path), 0)
        string = store.get_value(store.get_iter(path), 1)
        if self.activation_action:
            self.activation_action(task)

    def task_displayed(self, task):
        """Returns a flag indicating whether the referenced task should be displayed."""
        result = False
        if task.state != 'deleted':
            if task.has_runable_child():
                result = True
        return result

    def add_task_to_store(self, task, parent=None, sibling=None):
        """Adds the specified task and it's children to the tree store."""
        iter = None
        if self.task_displayed(task):
            text = task.full_name() if config['todo'].get('full_name') else task.name
            iter = self.store.insert_before(parent, sibling, (task, text))
            task_list = task.children
            if config.get('sort_tasks'):
                task_list.sort()
            if config.get('submenus_on_top'):
                sub_menus = [child for child in task_list if child.children]
                elements = [child for child in task_list if not child.children]
                task_list = sub_menus + elements
            for child in task_list:
                self.add_task_to_store(child, iter)
        return iter

class Find_Task_Dialog(Task_List_Dialog):
    """Dialog to diplay when no current activity selected."""

    def __init__(self, session):
        """Create an instance of the dialog."""
        Task_List_Dialog.__init__(self, session, 'find_task_dialog')
        self.set_title('Find Task')
        self.activation_action = self.activate_cb
        
    def activate_cb(self, task):
        """Display the task dialog for the selected task."""
        Task_Dialog(self.session.task_manager, self.session, task)

class Start_Task_Dialog(Task_List_Dialog):
    """Dialog to diplay when no current activity selected."""

    def __init__(self, session):
        """Create an instance of the dialog."""
        Task_List_Dialog.__init__(self, session, 'start_task_dialog')
        self.set_title('Start Task')
        self.activation_action = self.activate_cb
        
    def activate_cb(self, task):
        """Display the task dialog for the selected task."""
        if task and task.is_runable():
            self.session.start_task(task)
            self.__class__.instance = None
            self.destroy()
        else:
            Task_Dialog(self.session.task_manager, self.session, start=True)
            self.__class__.instance = None
            self.destroy()

class Day(object):
    """Collection of data associated with timekeeping for a day."""
    __slots__ = ['start', 'end', 'hours', 'unbilled']
    
    def __init__(self, date):
        """Set default values."""
        self.start = date.replace(hour=int(config['day_start'][:2]), minute=int(config['day_start'][3:]))
        day_duration = config['hours_per_week'] / config['days_per_week'] + config['lunch_duration']
        self.end = self.start + datetime.timedelta(hours = day_duration)
        self.hours = 0.0
        self.unbilled = 0.0
    
    def get_duration(self):
        """Adjust the end time based on updated hours."""
        return get_decimal_hours(self.end - self.start)

    def set_end_time(self):
        """Adjust the end time based on updated hours."""
        total = self.hours + self.unbilled
        self.end = self.start + datetime.timedelta(hours=total)

    def changed_finish(self):
        """Adjust the hours based on an updated end time."""
        new_total = get_decimal_hours(self.end - self.start)
        old_total = self.hours + self.unbilled
        if new_total > old_total:
            self.hours += new_total - old_total
        else:
            if old_total - new_total > self.unbilled:
                self.unbilled = 0.0
                self.hours = new_total
            else:
                self.unbilled -= old_total - new_total

class Day_View(object):
    """Gui elements displaying day class."""
    __slots__ = ['day', 'start', 'end', 'hours', 'unbilled',
                 'hours_changed', 'unbilled_changed', 'start_changed', 'end_changed']
    
    def __init__(self, day):
        """Store reference to day object and create widgets."""
        now = datetime.datetime.now()        
        self.day = day
        self.hours = gtk.Entry(6)
        #self.hours.set_override_mode(True)
        self.hours.set_width_chars(6)
        self.hours.set_text("%.2f" % (day.hours))
        if now > day.end:
            self.hours.set_editable(False)
            self.hours.set_sensitive(False)
        else:
            self.hours_changed = self.hours.connect('changed', self.hours_changed_cb)
        
        self.unbilled = gtk.Entry(6)
        #self.unbilled.set_override_mode(True)
        self.unbilled.set_width_chars(6)
        self.unbilled.set_text("%.2f" % (day.unbilled))
        if now > day.end:
            self.unbilled.set_editable(False)
            self.unbilled.set_sensitive(False)
        else:
            self.unbilled_changed = self.unbilled.connect('changed', self.unbilled_changed_cb)
        
        self.start = gtk.Entry(8)
        #self.start.set_override_mode(True)
        self.start.set_width_chars(8)
        self.start.set_text(day.start.strftime(config['time_format']))
        if now > day.start:
            self.start.set_editable(False)
            self.start.set_sensitive(False)
        else:
            self.start_changed = self.start.connect('changed', self.start_changed_cb)            

        self.end = gtk.Entry(8)
        #self.end.set_override_mode(True)
        self.end.set_width_chars(8)
        self.end.set_text(day.end.strftime(config['time_format']))
        if now > day.end:
            self.end.set_editable(False)
            self.end.set_sensitive(False)
        else:
            self.end_changed = self.end.connect('changed', self.end_changed_cb)

    def start_changed_cb(self, entry):
        """Updates the end time based on new start time."""
        self.day.start = string_to_time(self.start.get_text())
        self.day.set_end_time()
        self.update_end()

    def hours_changed_cb(self, entry):
        """Updates the end time based on new hours."""
        self.day.hours = float(self.hours.get_text())
        self.day.set_end_time()
        self.update_end()

    def unbilled_changed_cb(self, entry):
        """Updates the end time based on new hours."""
        self.day.unbilled = float(self.unbilled.get_text())
        self.day.set_end_time()
        self.update_end()

    def end_changed_cb(self, entry):
        """Updates the hours based on new end time."""
        self.day.end = string_to_time(self.end.get_text())
        self.day.changed_finish()
        self.update_hours()

    def update_end(self):
        """Sets the end time without triggering the change handler."""
        self.end.handler_block(self.end_changed)
        self.end.set_text(self.day.end.strftime(config['time_format']))
        self.end.handler_unblock(self.end_changed)

    def update_hours(self):
        """Sets the hours without triggering the change handler."""
        self.hours.handler_block(self.hours_changed)
        self.unbilled.handler_block(self.unbilled_changed)

        self.hours.set_text("%.2f" % (self.day.hours))
        self.unbilled.set_text("%.2f" % (self.day.unbilled))

        self.hours.handler_unblock(self.hours_changed)
        self.unbilled.handler_unblock(self.unbilled_changed)

class Timekeeping_Dialog(gtk.Dialog, Managed_Dialog):
    """Dialog to show timekeeping."""
    
    def __init__(self, session):
        """Create a simple dialog to show timekeeping."""
        gtk.Dialog.__init__(self, 'Timekeeping', None, 0,
                            (gtk.STOCK_OK, gtk.RESPONSE_ACCEPT))
        Managed_Dialog.__init__(self, "timekeeping_dialog")
        self.set_title('Timekeeping')
        self.session = session
        now = datetime.datetime.now()
        table = gtk.Table(3, 9)
        label = gtk.Label("Start")
        table.attach(label, 0, 1, 1, 2, gtk.FILL, gtk.FILL)
        label = gtk.Label("Hours")
        table.attach(label, 0, 1, 2, 3, gtk.FILL, gtk.FILL)
        label = gtk.Label("Unbilled")
        table.attach(label, 0, 1, 3, 4, gtk.FILL, gtk.FILL)
        label = gtk.Label("Finish")
        table.attach(label, 0, 1, 4, 5, gtk.FILL, gtk.FILL)

        tasks = [{}] * 7
        timekeeping = []
        for day in range(0, 7):
            date = self.session.week_start + datetime.timedelta(day)
            timekeeping.append(Day(date))
            # For days before today set the start to end of day and end to start of day.
            # These will be updated as tasks are processed.
            if now > timekeeping[day].start:
                timekeeping[day].start = self.session.week_start + datetime.timedelta(day + 1)
            if now > timekeeping[day].end:
                timekeeping[day].end = self.session.week_start + datetime.timedelta(day)

        activities = [act for act in session.activities if act.task.is_billable()]
        for activity in session.activities:
            start = activity.start_event.time
            end = start.replace(hour=0, minute=0, second=0, microsecond=0) + datetime.timedelta(1)
            # Allow for activities spanning multiple days
            while end < activity.end_event.time:
                duration = get_decimal_hours(end - start)
                day = start.weekday()
                if activity.task.is_billable():
                    timekeeping[day].hours += duration
                    if task in tasks[day]:
                        tasks[day][task] += duration
                    else:
                        tasks[day][task] = duration
                else:
                    timekeeping[day].unbilled += duration
                task = activity.task.full_name()
                if start < timekeeping[day].start:
                    timekeeping[day].start = start
                if end > timekeeping[day].end:
                    timekeeping[day].end = end
                start = end
                end += datetime.timedelta(1)
            day = start.weekday()
            end = activity.end_event.time
            duration = get_decimal_hours(end - start)
            # Convert all day task to hours per day
            if duration == full_day:
                duration = config['hours_per_week'] / config['days_per_week']
            if activity.task.is_billable():
                timekeeping[day].hours += duration
                task = activity.task.full_name()
                if task in tasks[day]:
                    tasks[day][task] += duration
                else:
                    tasks[day][task] = duration
            else:
                timekeeping[day].unbilled += duration
            if start < timekeeping[day].start:
                timekeeping[day].start = start
            if end > timekeeping[day].end:
                timekeeping[day].end = end

        spread_hours = 0.0
        if sum([day.hours for day in timekeeping if day.start.weekday() in weekdays]) < config['hours_per_week']:
            today = now.weekday()
            if today < 5:
                if today > 0:
                    allocated = sum([day.hours for day in timekeeping[:today]])
                    spread_hours = (config['hours_per_week'] - allocated) / (5.0 - today)
                else:
                    spread_hours = config['hours_per_week'] / 5.0

        total = [0.0, 0.0]
        for day in range(0, 7):
            if day > 4:
                spread_hours = 0.0
            start = self.session.week_start + datetime.timedelta(day)
            
            label = gtk.Label(start.strftime('%A'))
            label.set_width_chars(10)
            table.attach(label, day + 1, day + 2, 0, 1, gtk.FILL, gtk.FILL)

            if now > start + datetime.timedelta(1):
                timekeeping[day].unbilled = timekeeping[day].get_duration() - timekeeping[day].hours
            elif now > start:
                timekeeping[day].hours = spread_hours
                timekeeping[day].set_end_time()
            else:
                timekeeping[day].hours = spread_hours
                if timekeeping[day].unbilled == 0.0 and day < 5:
                    timekeeping[day].unbilled = 1.0
                timekeeping[day].set_end_time()
                
            day_ui = Day_View(timekeeping[day])
            total[0] += timekeeping[day].hours
            total[1] += timekeeping[day].unbilled
            table.attach(day_ui.start, day + 1, day + 2, 1, 2, gtk.FILL, gtk.FILL)
            table.attach(day_ui.hours, day + 1, day + 2, 2, 3, gtk.FILL, gtk.FILL)
            table.attach(day_ui.unbilled, day + 1, day + 2, 3, 4, gtk.FILL, gtk.FILL)
            table.attach(day_ui.end, day + 1, day + 2, 4, 5, gtk.FILL, gtk.FILL)

        table.attach(gtk.Label("Total"), day + 2, day + 3, 0, 1, gtk.FILL, gtk.FILL)

        entry = gtk.Entry(6)
        entry.set_width_chars(6)
        entry.set_text("%.2f" % (total[0]))
        table.attach(entry, day + 2, day + 3, 2, 3, gtk.FILL, gtk.FILL)
        entry = gtk.Entry(6)
        entry.set_width_chars(6)
        entry.set_text("%.2f" % (total[1]))
        table.attach(entry, day + 2, day + 3, 3, 4, gtk.FILL, gtk.FILL)

        self.vbox.pack_start(table)
        self.connect('response', self.response_cb)
        self.show_all()

    def finish_changed(self, day):
        """Adjust the hours based on the new finish time."""
        new_total = get_decimal_hours(day.end_time - day.start_time)

    def response_cb(self, widget, response):
        """Handle a response by closing the window."""
        self.destroy() 

class Reminder_Dialog(gtk.Dialog):
    """Dialog to diplay when current activity crosses existing event."""

    __gsignals__ = {'changed' : (gobject.SIGNAL_RUN_LAST, gobject.TYPE_NONE, (gobject.TYPE_INT,))}

    last_action = 1
    reminder = None
    allow_updates = True

    @classmethod
    def add_events(cls, session, events):
        """Add the events to the reminder if one exists, else create one."""
        new_reminder = None
        if cls.reminder:
            cls.reminder.add(events)
        else:
            cls.reminder = cls(session, events)
            new_reminder = cls.reminder
        return new_reminder

    def __init__(self, session, events):
        """Create and displays the reminder dialog."""
        gtk.Dialog.__init__(self, 'Reminder')
        self.tasks = {}
        self.events = events
        self.session = session
        self.connect('response', self.response_cb)
        self.store = gtk.ListStore(object, gobject.TYPE_STRING, gobject.TYPE_STRING)
        self.view = gtk.TreeView(self.store)
        self.view.get_selection().set_mode(gtk.SELECTION_SINGLE)

        self.options = gtk.ListStore(gobject.TYPE_STRING)
        for option in ['Start When Due', 'Start Now', 'Delay Activity', 'Ignore Activity', 'Delete Activity']:
            self.options.append([option])

        self.end_options = gtk.ListStore(gobject.TYPE_STRING)
        for option in ['End When Due', 'End Now', 'Keep Going']:
            self.end_options.append([option])

        self.view.set_headers_visible(False)
        text_renderer = gtk.CellRendererText()
        tree_column = gtk.TreeViewColumn(None, text_renderer, text = 1)
        tree_column.set_sizing(gtk.TREE_VIEW_COLUMN_AUTOSIZE)
        self.view.append_column(tree_column)

        combo_renderer = gtk.CellRendererCombo()
        combo_renderer.set_property('editable', True)
        combo_renderer.set_property('has-entry', False)
        combo_renderer.set_property('model', self.options)
        combo_renderer.set_property('text-column', 0)
        combo_renderer.connect('edited', self.action_edited_cb)
        combo_renderer.connect('editing-started', self.editing_cb)
        combo_renderer.connect('editing-canceled', self.editing_canceled_cb)
        #combo_renderer.connect('changed', self.changed_cb)
        tree_column = gtk.TreeViewColumn(None, combo_renderer, text = 2)
        tree_column.set_sizing(gtk.TREE_VIEW_COLUMN_AUTOSIZE)
        self.view.append_column(tree_column)
        self.vbox.pack_start(self.view)
        self.update_text()

        remind = gtk.HBox()
        sleep_button = gtk.Button('Remind me in')
        sleep_button.connect('clicked', self.sleep_cb)
        self.sleep_time = gtk.SpinButton()
        adjust = self.sleep_time.get_adjustment()
        adjust.set_all(5, 1, 60, 1, 5, 0)
        label = gtk.Label('mins')
        remind.pack_start(sleep_button, False, False)
        remind.pack_start(self.sleep_time, False, False)
        remind.pack_start(label, False, False)

        self.action_area.pack_start(remind, False, False)

        self.cancel_button = gtk.Button(stock = gtk.STOCK_CANCEL)
        self.add_action_widget(self.cancel_button, gtk.RESPONSE_CANCEL)
        self.ok_button = gtk.Button(stock = gtk.STOCK_OK)
        self.add_action_widget(self.ok_button, gtk.RESPONSE_OK)

        self.show_all()

    def editing_cb(self, cell_renderer, editable, path):
        """Disable updates while editing cells."""
        self.allow_updates = False

    def editing_canceled_cb(self, cell_renderer):
        """Re-enables updates."""
        self.allow_updates = True

    def sleep_cb(self, button):
        """Hides the dialog for the specified time."""
        time = int(self.sleep_time.get_text()) * 60000
        self.allow_updates = False
        self.hide()
        self.timeout = gobject.timeout_add(time, self.redisplay_cb)

    def redisplay_cb(self):
        """Unhide the window and update it's contents."""
        self.allow_updates = True
        self.show_all()
        self.present()
        self.update()
        return False

    def action_edited_cb(self, renderer, path, action):
        """Responds to the task state being edited."""
        task_iter = self.store.get_iter(path)
        task = self.store.get_value(task_iter, 0)
        self.update_action(task, action)
        self.allow_updates = True
        self.update_text()

    def changed_cb(self, renderer, path, action_iter):
        """Responds to the task state being changed."""
        task_iter = self.store.get_iter(path)
        task = self.store.get_value(task_iter, 0)
        action = self.options.get_value(action_iter, 0)
        self.update_action(task, action)
        self.allow_updates = True
        self.update_text()

    def update_action(self, task, action):
        """Updates the action of the specified task and others as necessary."""
        activity = self.tasks[task]
        activity.action = action
        event = activity.start_event
        if action == 'Start Now' or action == 'Start When Due':
            for other_activity in event.subsequent_activity:
                if other_activity != activity:
                    other_activity.action = 'Ignore'
            activity.action = action
        elif action == 'Delay Activity':
            for other_activity in event.subsequent_activity:
                other_activity.action = 'Delay Activity'

    def add(self, events):
        """Add the task in the supplied events to the reminder."""
        self.events.extend(events)
        self.update_text()

    def redraw(self):
        """Updates the dialog."""
        self.update_text()

    def update(self):
        """Updates the dialog."""
        self.update_text()

    def update_text(self):
        """Updates the dialog text based on the status."""
        if not self.allow_updates:
            return
        self.tasks = {}
        for event in self.events:
            event.remind = False
            for activity in event.subsequent_activity:
                if activity.task:
                    task = activity.task
                    if task not in self.tasks or event.time < self.tasks[task].start_event.time:
                        self.tasks[task] = activity
        if len(self.tasks) < 1:
            self.destroy()
            return
        self.store.clear()
        for task in self.tasks.keys():
            time = self.tasks[task].start_event.time
            now = datetime.datetime.now()
            if time > now:
                time_delta = time - now
                text = '%s is due to start in %s' % (task.name, timedelta_to_string(time_delta))
            else:
                time_delta = now - time
                text = '%s is %s overdue' % (task.name, timedelta_to_string(time_delta))
            if self.tasks[task].action:
                action_text = self.tasks[task].action
            else:
                action_text = 'Ignore Activity'
            self.store.append((task, text, action_text))

    def response_cb(self, dialog, response_id):
        """Processes the user response to current activity overwriting an event."""
        if response_id == gtk.RESPONSE_CANCEL:
            for task in self.tasks.keys():
                activity = self.tasks[task]
                activity.action = None
        else:
            for task in self.tasks.keys():
                activity = self.tasks[task]
                if activity.action == 'Delete Activity':
                    del self.tasks[task]
                    activity.delete()
                elif activity.action == 'Ignore Activity' or not activity.action:
                    activity.allocation = 0
                elif activity.action == 'Delay Activity':
                    activity.start_event.action = 'Delay'
                    activity.start_event.remind = True
                elif activity.action == 'Start Now':
                    event = self.session.new_event()
                    activity.change_start(event)
                    self.current_activity = activity
                    del activity.action
                elif activity.action == 'Start When Due':
                    activity.start_event.action = activity

        self.emit('changed', 0)
        self.destroy()

class App_Gui(gtk.StatusIcon):
    """Applications main GUI class"""

    windows = []
    dialogs = []

    def __init__(self):
        """Initialise the task application GUI"""
        gtk.StatusIcon.__init__(self)
        self.set_from_file(config.get('image_root') + '22x22/apps/preferences-system-time.png')
        self.set_tooltip('Time booking system')
        self.connect('activate', self.activate_cb)
        self.connect('popup-menu', self.menu_cb)
        self.task_manager = Task_Manager(self.update)
        if config.get('custom_session_file'):
            session_file = config['custom_session_file']
            del config['custom_session_file']
        else:
            session_file = None
        self.session = Activity_Manager(self.task_manager,
                                        self.update,
                                        filename = session_file)
        self.task_list = None

        gtk.notebook_set_window_creation_hook(self.new_window_cb, None)

        self.update_task_menu()
        self.update_options_menu()

        if config.get('startup_task') == 'previous':
            if not self.session.current_activity:
                if self.session.previous_task:
                    self.session.start_task(self.session.previous_task)
                else:
                    Start_Task_Dialog.show(self.session)
                    Start_Task_Dialog.instance.connect('changed', self.timesheet_changed_cb)
                    #self.task_menu.popup(None, None, None, 1, gtk.get_current_event_time())
            elif (self.session.current_activity.end_event and 
                  self.session.current_activity.end_event.get_id() != 0):
                task = self.session.current_activity.task
                self.session.current_activity = None
                self.session.start_task(task)

        if config['timesheet'].get('show'):
            self.timesheet_cb(self)

        if config['todo'].get('show'):
            self.task_list_cb(self)

        if config['notepad'].get('show'):
            self.notepad_cb(self)

        if config.get('startup_task') == 'ask' and not self.session.current_activity:
            Start_Task_Dialog.show(self.session)
            Start_Task_Dialog.instance.connect('changed', self.timesheet_changed_cb)

        #if config['outlook_import']:
        #    self.session.outlook_import_cb(self)

        # Add a periodic update to keep display up to date.
        self.timeout = gobject.timeout_add(10000, self.update)

    def set_status(self):
        """Update the tool tip."""
        if self.session.current_activity:
            text = '\n%s active' % (self.session.current_activity.task.name)
        elif self.session.previous_task:
            text = '\n%s paused' % (self.session.previous_task.name)
        else:
            text = ''
        self.set_tooltip("Time booking system%s" % (text))

    def new_window_cb(self, source, page, x, y, user_data):
        """Creates a new window for a dropped notebook tab."""
        widget = Timesheet_Editor(self.session, tab = page)
        new_window = self.new_window(widget)
        new_window.move(x, y)
        return new_window.tabs

    def new_window(self, window):
        """Creates a new window."""
        window.connect('delete-event', self.window_deleted_cb)
        window.connect('changed', self.timesheet_changed_cb)
        self.windows.append(window)
        return window

    def window_deleted_cb(self, window, event):
        """Removes the deleted window from the window list."""
        self.windows.remove(window)

    def timesheet_changed_cb(self, widget, data):
        """Update the GUI in response to a change on a timesheet widget."""
        self.redraw()

    def task_dialog_cb(self, item, task):
        """Displays the task dialog for the specified task."""
        Task_Dialog(self.task_manager, self.session, task)

    def update(self):
        """Updates all displayed windows."""
        self.set_status()
        if self.session.current_activity:
            text = self.session.current_activity.task.name
            self.set_tooltip("Time booking system\n%s" % (text))
            if Start_Task_Dialog.instance:
                Start_Task_Dialog.instance.destroy()
        else:
            self.set_tooltip("Time booking system")
            if Start_Task_Dialog.instance == -1:
                Start_Task_Dialog.instance = None
            Start_Task_Dialog.show(self.session)
            Start_Task_Dialog.instance.connect('changed', self.timesheet_changed_cb)

        for window in self.windows:
            window.update()
            window.queue_draw()
        for dialog in self.dialogs:
            dialog.update()
        if self.session.current_activity:
            self.session.current_activity.update_duration()
        events = self.session.check_upcoming_events()
        events = [event for event in events if event.subsequent_activity and event.remind]
        if len(events) > 0:
            self.create_reminder(events)
        return True

    def create_reminder(self, events):
        """Creates a reminder for the given events adding to current dialog if present."""
        new_dialog = Reminder_Dialog.add_events(self.session, events)
        if new_dialog:
            new_dialog.connect('unrealize', self.dialog_unmapped_cb)
            new_dialog.connect('changed', self.dialog_changed_cb)
            self.dialogs.append(new_dialog)

    def dialog_unmapped_cb(self, dialog):
        """Removes the deleted dialog from the dialog list."""
        self.dialogs.remove(dialog)
        Reminder_Dialog.reminder = None

    def dialog_changed_cb(self, widget, data):
        """Updates the other GUI elements in response to a dialog change."""
        self.redraw()

    def redraw(self):
        """Redraws all displayed windows."""
        self.set_status()
        for window in self.windows:
            window.redraw()
        for dialog in self.dialogs:
            dialog.redraw()
        events = self.session.check_upcoming_events()
        events = [event for event in events if event.subsequent_activity and event.remind]
        if len(events) > 0:
            self.create_reminder(events)

    def get_task_menu(self):
        """Returns the task menu definition."""
        menu = [('New Task', self.new_task_cb)]
        task_menu = self.task_manager.create_menu_definition(self.session.activate_cb,
                                                             config.get('hierarchical_list'))
        resume_menu = []
        if self.session.previous_task:
            if not self.session.current_activity or self.session.current_activity.task != self.session.previous_task:
                resume_menu.append(('Resume %s' % self.session.previous_task.name, self.resume_cb))
        if self.session.current_activity:
            resume_menu.append(('Pause %s' % (self.session.current_activity.task.name), self.pause_cb))

        if task_menu:
            menu.append(('---', '---'))
            menu.extend(task_menu)
        if resume_menu:
            menu.append(('---', '---'))
            menu.extend(resume_menu)
        return menu

    def update_task_menu(self):
        """Updates the task menu."""
        menu_definition = self.get_task_menu()
        self.task_menu = build_menu(menu_definition)
        self.task_menu.connect('leave-notify-event', self.leave_event_cb)

    def get_options_menu(self):
        """Returns the options menu definition."""
        task_menu = self.task_manager.create_menu_definition(self.task_dialog_cb,
                                                             config.get('hierarchical_list'),
                                                             all_tasks = True)
        options_menu = [('>>>Tasks', task_menu),
                        ('Options', self.options_cb),
                        ('Show Timesheet', self.timesheet_cb),
                        ('Show Timekeeping', self.timekeeping_cb),
                        ('Show Notepad', self.notepad_cb),
                        ('Show To Do List', self.task_list_cb)]
        if self.windows:
            options_menu.append(('Close Windows', self.close_windows_cb))

        options_menu.append(('Exit', self.exit_cb))

        return options_menu

    def update_options_menu(self):
        """Updates the options menu."""
        self.options_menu = build_menu(self.get_options_menu(), False)
        self.options_menu.add_events(gtk.gdk.FOCUS_CHANGE_MASK)
        self.options_menu.connect('leave-notify-event', self.leave_event_cb)

    def leave_event_cb(self, widget, event=None):
        """Removes the menu when the mouse leaves it."""
        if event.detail == gtk.gdk.NOTIFY_UNKNOWN:
            widget.popdown()

    def activate_cb(self, status_icon):
        """Displays the task selection menu."""
        self.update_task_menu()
        self.task_menu.popup(None, None, None, 1, gtk.get_current_event_time())

    def menu_cb(self, status_icon, button, activate_time):
        """Displays the options menu."""
        self.update_options_menu()
        self.options_menu.popup(None, None, None, button, activate_time)

    def new_task_cb(self, widget):
        """Displays the new task dialog."""
        Task_Dialog(self.task_manager, self.session)
        self.update()

    def pause_cb(self, widget):
        """"Pauses the current activity."""
        self.session.pause()
        self.update()

    def resume_cb(self, widget):
        """Resumes the previous activity."""
        self.session.resume()
        self.redraw()

    def close_windows_cb(self, widget):
        """Close all the windows currently displayed."""
        for window in self.windows:
            window.destroy()
        self.windows = []

    def exit_cb(self, widget):
        """Quits the application."""
        self.close_windows_cb(widget)
        if config.get('pause_on_exit') and self.session.current_activity:
            self.session.pause()
        self.session.save()
        gtk.main_quit()

    def options_cb(self, widget):
        """Displays the options gui."""
        Options_Dialog()

    def task_list_cb(self, widget):
        """Display the todo list window."""
        window = Task_List(self.session)
        window.connect('show-task-selection', self.activate_cb)
        window.connect('timesheet-request', self.timesheet_cb)
        window.connect('find-task-request', self.find_task_cb)
        window.connect('notepad-request', self.notepad_cb)
        window.connect('exit-request', self.exit_cb)
        self.new_window(window)

    def timekeeping_cb(self, widget):
        """Displays the timekeeping dialog."""
        Timekeeping_Dialog(self.session)

    def notepad_cb(self, widget):
        """Displays the notepad."""
        window = Notepad.instance(self.session)
        self.new_window(window)        

    def timesheet_cb(self, widget):
        """Display the timesheet window."""
        window = Timesheet_Editor(self.session)
        self.new_window(window)        

    def find_task_cb(self, widget):
        """Display the find task window."""
        Find_Task_Dialog.show(self.session)

def initialise():
    """Sets up the options parser and returs a configuration object based on the options."""
    parser = OptionParser()
    parser.add_option("-d", "--data_dir", default = defaults['data_dir'],
                      dest = "data_dir", help = "Directory for data files",
                      metavar = "FILE")

    parser.add_option("--session_file", dest = "session_file",
                      help = "load specific session file", default = '',
                      metavar = "FILE")

    parser.add_option("-f", "--file", dest = "config_file",
                      help = "load configuration file", default = defaults['config_file'],
                      metavar = "FILE")

    parser.add_option("-s", "--session", dest = "initial_session",
                      help = "Session to load")

    parser.add_option("-t", "--timesheet", action="store_true", default=False,
                      dest = "show_timesheet", help = "start timesheet view")

    parser.add_option("-i", "--images", default= defaults['image_root'],
                      dest = "image_root", help = "location of the png images")

    parser.add_option("-v", "--verbose", action="store_true", default=False,
                      dest = "verbose", help = "display verbose output")

    (options, args) = parser.parse_args()
    return Configuration(options)


##############################################################################################################
# If the program is run directly or passed as an argument to the python
# interpreter then create a code edit instance and show it
if __name__ == "__main__":
    os.chdir(os.path.split(sys.argv[0])[0])
    user = User()

    config = initialise()

    now = datetime.datetime.now()
    start_time = now.replace(hour = 0,
                             minute = 0,
                             second = 0,
                             microsecond = 0)
    end_time =  now.replace(hour = 23,
                            minute = 59,
                            second = 59,
                            microsecond = 0)
    full_day = get_decimal_hours(end_time - start_time)

    root_path = os.path.split(sys.argv[0])[0]
    if root_path == '.':
        root_path = os.getcwd()

    # Dummy button to ensure gtk-button-images property has been created
    gtk.Button()
    gtk.settings_get_default().set_property('gtk-button-images', True)

    window = App_Gui()
    try:
        gtk.main()
    finally:
        window.set_visible(False)
        config.save()
