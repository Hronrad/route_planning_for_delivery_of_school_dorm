from lib2to3.pgen2.token import GREATER

from matplotlib.pyplot import text
import kivy
from kivy.app import App
from kivy.uix.label import Label
from kivy.uix.gridlayout import GridLayout
from kivy.uix.stacklayout import StackLayout
from kivy.uix.textinput import TextInput
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.widget import Widget
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.uix.floatlayout import FloatLayout
from kivy.uix.popup import Popup
from kivy.properties import ObjectProperty
from kivy.properties import StringProperty
from kivy.properties import BooleanProperty
from kivy.lang import Builder
from kivy.uix.recycleview import RecycleView
from kivy.uix.recycleview.views import RecycleDataViewBehavior
from kivy.uix.recycleboxlayout import RecycleBoxLayout
from kivy.uix.behaviors import FocusBehavior
from kivy.uix.recycleview.layout import LayoutSelectionBehavior

import os
import openpyxl
from numpy import zeros_like

kivy.require('2.1.0')


romlist = []
numlist = []
checklist = []
Nroom, Nnum, Len, Cur = 0, 0, 0, 0

popupWindow = Popup()


class TestApp(App):
    def pop(self):
        global romlist, numlist, checklist, Nroom, Nnum, Len, Cur
        popupWindow.dismiss()
        show_popup()

    def setValue(self):
        global Len
        RV.alter_data(RV)
        return Len

    def dismis(self):
        popupWindow.dismiss()

    def update(self):
        global romlist, numlist, checklist, Nroom, Nnum, Len, Cur
        Cur = romlist.index(Nroom)
        checklist[Cur] = 1
        App.get_running_app().rv = self
        RV.alter_data(RV)
        if(Cur+1 < Len):
            Cur = Cur+1
            Nroom = romlist[Cur]
            Nnum = numlist[Cur]
        else:
            Cur = 999

    def getValue(self, str1):
        global romlist, numlist, checklist, Nroom, Nnum, Len, Cur
        print(eval(str1))
        return eval(str1)

    def build(self):

        return kv


def getValue1(str1):
    global romlist, numlist, checklist, Nroom, Nnum, Len, Cur
    print(eval(str1))
    return eval(str1)


class LoginWindow(Screen):
    pass


class P(FloatLayout):
    pass


def show_popup():
    global romlist, numlist, checklist, Nroom, Nnum, Len, Cur, popupWindow
    show = P()
    popupWindow = Popup(
        title=f"Room:{Nroom}    Num:{Nnum}————{Cur+1}/{Len}", content=show, size_hint=(None, None), size=(400, 200))

    popupWindow.open()


class FileChooserWindow(Screen):
    def open(self, path, filename):
        global romlist, numlist, checklist, Nroom, Nnum, Len, Cur

        romlist, numlist = loadxlfile(os.path.join(path, filename[0]))
        Len = len(numlist)
        checklist = zeros_like(numlist)
        Nroom = romlist[0]
        Nnum = numlist[0]


class SelectableRecycleBoxLayout(FocusBehavior, LayoutSelectionBehavior,
                                 RecycleBoxLayout):
    pass


class RV(RecycleView, App):
    def __init__(self, **kwargs):
        global romlist, numlist, checklist, Nroom, Nnum, Len, Cur
        super(RV, self).__init__(**kwargs)
        self.data = [{'text': str(i)}
                     for i in range(50)]

    def alter_data(self):
        global romlist, numlist, checklist, Nroom, Nnum, Len, Cur
        self.data = [{'text': f'{romlist[i]}——{numlist[i]}'}
                     for i in range(Len)]


class ListWindow(Screen):
    def refresh_recycleview(self):
        global romlist, numlist, checklist, Nroom, Nnum, Len, Cur
        self.ids.rv_id.data = [{'text': str(i)}
                               for i in range(50)]


class UsageWindow(Screen):
    pass


class WindowManager(ScreenManager):
    pass


class SelectableLabel(RecycleDataViewBehavior, Label):
    ''' Add selection support to the Label '''
    index = None
    text = StringProperty()
    generated_state_text = StringProperty()
    active = BooleanProperty()
    selected = BooleanProperty(False)
    selectable = BooleanProperty(True)

    def refresh_view_attrs(self, rv, index, data):
        ''' Catch and handle the view changes '''
        self.index = index
        if data['text']:
            self.generated_state_text = '1'
        return super(SelectableLabel, self).refresh_view_attrs(
            rv, index, data)

    def on_touch_down(self, touch):
        ''' Add selection on touch down '''
        if super(SelectableLabel, self).on_touch_down(touch):
            return True
        if self.collide_point(*touch.pos) and self.selectable:
            return self.parent.select_with_touch(self.index, touch)

    def apply_selection(self, rv, index, is_selected):
        ''' Respond to the selection of items in the view. '''
        self.selected = is_selected
        if is_selected:
            print("selection changed to {0}".format(rv.data[index]))
        else:
            print("selection removed for {0}".format(rv.data[index]))


kv = Builder.load_string(
    '''
    #:import Button kivy.uix.button.Button
#:set lenth 1
WindowManager:
    LoginWindow:
    FileChooserWindow:
    ListWindow:
    UsageWindow:

<SelectableLabel>:
    # Draw a background to indicate selection
    canvas.before:
        Color:
            rgba: (.0, 0.9, .1, .3) if self.selected else (0, 0, 0, 1)
        Rectangle:
            pos: self.pos
            size: self.size




# <ListWindow>:
#     name:"ListWindow"
#     GridLayout:
#         id:GridLayout3
#         rows:2
#         viewclass: 'SelectableLabel'
#         SelectableRecycleBoxLayout:
#             default_size: None, dp(56)
#             default_size_hint: 1, None
#             size_hint_y: None
#             height: self.minimum_height
#             orientation: 'vertical'
#             multiselect: True
#             touch_multiselect: True
#         GridLayout:
#             id:GridLayout31
#             size_hint_y:None
#             height:40
#             cols:2
#             Button:
#                 text:"LIST"
#                 on_release:
#                     app.root.current='ListWindow'
                    
                
            
#             Button:
#                 text:"NAVIG"
#                 on_release:
#                     app.pop()
#                     # app.root.current='UsageWindow'
#                     # root.manager.transition.direction="left"

<LoginWindow>:
    name:"LoginWindow"

    GridLayout:
        id:GridLayout1
        cols:1
        size:root.width,root.height

        GridLayout:
            cols:2
            id:GridLayout11
            Label:
                text:"Name: "
            
            TextInput:
                id:name
                multiline:False
            
            Label:
                text:"Telephone: "
            
            TextInput:
                id:telephone
                multiline:False

            Label:
                text:"BuildingNumber: "
            
            TextInput:
                id:buildingnumber
                multiline:False
        
        Button:
            text:"Enter"
            height:200
            size_hint_y:None
            on_release:app.root.current="FileChooserWindow" if (name.text and telephone.text and buildingnumber.text) else "LoginWindow"

<FileChooserWindow>:
    name:"FileChooserWindow"
    FloatLayout:

        Label:
            text:"choose an excel file:"
            pos_hint:{"top":1.42}

        FileChooserListView:
            id:filechooser
            on_selection:root.open(filechooser.path,filechooser.selection);app.root.current="ListWindow"



<P>:
    Button:
        id:checkbutton1
        text:'NEXT'
        size_hint:.3,.3
        pos_hint:{'center_x':.5,'center_y':.3}
        on_press:
            app.update()
            lenth=app.getValue('Len')
            if app.getValue('Cur')<app.getValue('Len'):app.pop() 
            if app.getValue('Cur')==999:app.dismis()



# <UsageWindow>:
#     name:"UsageWindow"
#     GridLayout:
#         id:GridLayout41
#         rows:2
#         FloatLayout:
#             id:FloatLayout41
#             Label:
#                 id:room
#                 text:app.getValue('Nroom')
#                 pos_hint:{"y":.42}
            
#             Label:
#                 id:num
#                 text:app.getValue('Nnum')
#                 pos_hint:{"y":.32}

#             Button:
#                 id:checkbutton
#                 text:"START"
#                 size_hint:.3,.3
#                 pos_hint:{'center_x':.5,'center_y':.3}
#                 on_press:app.update()

#         GridLayout:
#             id:GridLayout42
#             size_hint_y:None
#             height:40
#             cols:2
#             Button:
#                 text:"LIST"
#                 on_release:
#                     app.root.current="ListWindow"
#                     root.manager.transition.direction="right"
#             Button:
#                 text:"NAVIG"

<ListWindow>:
    on_pre_enter:self.refresh_recycleview;
    name:"ListWindow"
    BoxLayout:
        orientation:'horizontal'
        RV:
            id:rv_id
            viewclass: 'SelectableLabel'
            
            SelectableRecycleBoxLayout:
                default_size: None, dp(36)
                default_size_hint: 1, None
                size_hint_y: None
                height: self.minimum_height
                orientation: 'vertical'
                multiselect: True
                touch_multiselect: True

    GridLayout:
        id:GridLayout31
        size_hint_y:None
        height:40
        cols:2
        Button:
            text:"LIST"
            on_release:
                app.root.current='ListWindow'

        Button:
            text:"NAVIG"
            on_release:
                app.pop()
                # app.root.current='UsageWindow'
                # root.manager.transition.direction="left"
                
'''
)
if __name__ == '__main__':
    TestApp().run()


def loadxlfile(xlfile):
    wb = openpyxl.load_workbook(xlfile)
    sheet = wb['Sheet1']
    data = []  # 储存有效数据

    for cell_row in sheet:
        nobuilding = cell_row[0].value
        if (nobuilding == '1A' or nobuilding == '1a'):
            data.append({'A': cell_row})
        elif(nobuilding == '1B' or nobuilding == '1b'):
            data.append({'B': cell_row})

    tmp = tuple(data[0].values())
    total = tmp[0][3].value  # 应送总数
    dict = {}  # 门户字典
    for dics in data:
        for k, v in dics.items():
            section = k
            room = k+str(v[1].value)
            num = v[2].value
            if room in dict:
                dict[room] += num
            else:
                dict[room] = num

    romlist = []
    numlist = []
    sequence = ['A201', 'A202', 'A203', 'A204', 'A205', 'A206', 'A207', 'A208', 'A209', 'A210', 'A211', 'A212', 'A213', 'A214', 'A215', 'A216', 'A217', 'A218', 'A219', 'A220', 'A221', 'A222', 'A223', 'A224', 'A225', 'A226', 'A227', 'A228', 'A229', 'A230', 'A231', 'A232', 'A233', 'A234', 'A235', 'A236', 'A237', 'A238', 'A239', 'A240', 'A241', 'A242', 'A243', 'A244', 'A245', 'A246', 'A247', 'A248', 'A249', 'A250', 'A251', 'A252', 'A253', 'B253', 'B252', 'B251', 'B250', 'B249', 'B248', 'B247', 'B246', 'B245', 'B244', 'B243', 'B242', 'B241', 'B240', 'B239', 'B238', 'B237', 'B236', 'B235', 'B234', 'B233', 'B232', 'B231', 'B230', 'B229', 'B228', 'B227', 'B226', 'B225', 'B224', 'B223', 'B222', 'B221', 'B220', 'B219', 'B218', 'B217', 'B216', 'B215', 'B214', 'B213', 'B212', 'B211', 'B210', 'B209', 'B208', 'B207', 'B206', 'B205', 'B204', 'B203', 'B202', 'B201', 'B301', 'B302', 'B303', 'B304', 'B305', 'B306', 'B307', 'B308', 'B309', 'B310', 'B311', 'B312', 'B313', 'B314', 'B315', 'B316', 'B317', 'B318', 'B319', 'B320', 'B321', 'B322', 'B323', 'B324', 'B325', 'B326', 'B327', 'B328', 'B329', 'B330', 'B331', 'B332', 'B333', 'B334', 'B335', 'B336', 'B337', 'B338', 'B339', 'B340', 'B341', 'B342', 'B343', 'B344', 'B345', 'B346', 'B347', 'B348', 'B349', 'A349', 'A348', 'A347', 'A346', 'A345', 'A344', 'A343', 'A342', 'A341', 'A340', 'A339', 'A338', 'A337', 'A336', 'A335', 'A334', 'A333', 'A332', 'A331', 'A330', 'A329', 'A328', 'A327', 'A326', 'A325', 'A324', 'A323', 'A322', 'A321', 'A320', 'A319', 'A318', 'A317', 'A316', 'A315', 'A314', 'A313', 'A312', 'A311', 'A310', 'A309', 'A308', 'A307', 'A306', 'A305', 'A304', 'A303', 'A302', 'A301', 'A401', 'A402', 'A403', 'A404', 'A405', 'A406', 'A407', 'A408', 'A409', 'A410', 'A411', 'A412', 'A413', 'A414', 'A415', 'A416', 'A417', 'A418', 'A419', 'A420', 'A421', 'A422', 'A423', 'A424', 'A425', 'A426', 'A427', 'A428', 'A429', 'A430', 'A431', 'A432', 'A433', 'A434', 'A435', 'A436', 'A437', 'A438', 'A439', 'A440', 'A441', 'A442', 'A443', 'A444', 'A445', 'A446', 'A447', 'A448',
                'A449', 'B449', 'B448', 'B447', 'B446', 'B445', 'B444', 'B443', 'B442', 'B441', 'B440', 'B439', 'B438', 'B437', 'B436', 'B435', 'B434', 'B433', 'B432', 'B431', 'B430', 'B429', 'B428', 'B427', 'B426', 'B425', 'B424', 'B423', 'B422', 'B421', 'B420', 'B419', 'B418', 'B417', 'B416', 'B415', 'B414', 'B413', 'B412', 'B411', 'B410', 'B409', 'B408', 'B407', 'B406', 'B405', 'B404', 'B403', 'B402', 'B401', 'B501', 'B502', 'B503', 'B504', 'B505', 'B506', 'B507', 'B508', 'B509', 'B510', 'B511', 'B512', 'B513', 'B514', 'B515', 'B516', 'B517', 'B518', 'B519', 'B520', 'B521', 'B522', 'B523', 'B524', 'B525', 'B526', 'B527', 'B528', 'B529', 'B530', 'B531', 'B532', 'B533', 'B534', 'B535', 'B536', 'B537', 'B538', 'B539', 'B540', 'B541', 'B542', 'B543', 'B544', 'B545', 'B546', 'B547', 'B548', 'B549', 'A549', 'A548', 'A547', 'A546', 'A545', 'A544', 'A543', 'A542', 'A541', 'A540', 'A539', 'A538', 'A537', 'A536', 'A535', 'A534', 'A533', 'A532', 'A531', 'A530', 'A529', 'A528', 'A527', 'A526', 'A525', 'A524', 'A523', 'A522', 'A521', 'A520', 'A519', 'A518', 'A517', 'A516', 'A515', 'A514', 'A513', 'A512', 'A511', 'A510', 'A509', 'A508', 'A507', 'A506', 'A505', 'A504', 'A503', 'A502', 'A501', 'A601', 'A602', 'A603', 'A604', 'A605', 'A606', 'A607', 'A608', 'A609', 'A610', 'A611', 'A612', 'A613', 'A614', 'A615', 'A616', 'A617', 'A618', 'A619', 'A620', 'A621', 'A622', 'A623', 'A624', 'A625', 'A626', 'A627', 'A628', 'A629', 'A630', 'A631', 'A632', 'A633', 'A634', 'A635', 'A636', 'A637', 'A638', 'A639', 'A640', 'A641', 'A642', 'A643', 'A644', 'A645', 'A646', 'A647', 'A648', 'A649', 'A650', 'A651', 'A652', 'A653', 'B653', 'B652', 'B651', 'B650', 'B649', 'B648', 'B647', 'B646', 'B645', 'B644', 'B643', 'B642', 'B641', 'B640', 'B639', 'B638', 'B637', 'B636', 'B635', 'B634', 'B633', 'B632', 'B631', 'B630', 'B629', 'B628', 'B627', 'B626', 'B625', 'B624', 'B623', 'B622', 'B621', 'B620', 'B619', 'B618', 'B617', 'B616', 'B615', 'B614', 'B613', 'B612', 'B611', 'B610', 'B609', 'B608', 'B607', 'B606', 'B605', 'B604', 'B603', 'B602', 'B601']

    for item in sequence:
        if item in dict:
            romlist.append(item)
            numlist.append(dict[item])

    # lenth = len(romlist)
    # for i in range(lenth):
    #     print(romlist[i], numlist[i])

    checktotal = sum(numlist)
    print("check!")if checktotal == total else print("error!")
    return romlist, numlist
