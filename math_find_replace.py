#!/usr/bin/env python
# - * - mode: python; coding: utf-8 - * 

#  Copyright 2015 Andrey <ad.dycost@gmail.com>
#
#  Dialog box based on https://forum.openoffice.org/en/forum/viewtopic.php?f=45&t=56397
#
#
#  This program is free software; you can redistribute it and/or modify
#  it under the terms of the GNU General Public License as published by
#  the Free Software Foundation; either version 2 of the License, or
#  (at your option) any later version.
#  
#  This program is distributed in the hope that it will be useful,
#  but WITHOUT ANY WARRANTY; without even the implied warranty of
#  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#  GNU General Public License for more details.
#  
#  You should have received a copy of the GNU General Public License
#  along with this program; if not, write to the Free Software
#  Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston,
#  MA 02110-1301, USA.


# import the required modules
import uno
import unohelper

from com.sun.star.awt import WindowDescriptor
from com.sun.star.awt.PosSize import POSSIZE

class Main():

	def __init__(self):
		# get current document object
		self.doc = XSCRIPTCONTEXT.getDocument()
		self.component = XSCRIPTCONTEXT.getComponentContext()
		self.smgr = self.component.ServiceManager
		self.msg_1 = "Find:"
		self.msg_2 = "Replace to:"
		self.window_manage()


	def find_and_replace(self, symb_old, symb_new):
		'''Find symb_old symbol(s) and its replacement by symbol(s) symb_new'''
		# iterate through all draw objects
		for i in range(self.doc.DrawPage.getCount()):
			obj = self.doc.DrawPage.getByIndex(i)
			# checking for embedded object
			if  hasattr(obj, "EmbeddedObject"):
				obj_formula = obj.getEmbeddedObject()
				# checking for object is formula math
				if  hasattr(obj_formula, "Formula"):
					# get formula as string
					old_formula = obj_formula.Formula
					# make replace
					new_formula = old_formula.replace(symb_old, symb_new)
					obj_formula.Formula = new_formula
		return None


	def create_element(self, model_element, name_element, props_element):
		'''Insert UnoControl<srv>Model into given DialogControlModel oDM by given sName and properties dProps'''
		dialog_elem_model = self.dialog_win_model.createInstance("com.sun.star.awt.UnoControl"+ model_element +"Model")
		while props_element:
			prp = props_element.popitem()
			uno.invoke(dialog_elem_model, "setPropertyValue", (prp[0], prp[1]))
			# works with awt.UnoControlDialogElement only:
			dialog_elem_model.Name = name_element
		self.dialog_win_model.insertByName(name_element, dialog_elem_model)


	def window_manage(self):
		'''Show a dialog message box with the UNO based toolkit'''
		self.dialog_win_model = self.smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", self.component)

		# create the elements and set the properties
		self.dialog_win_model.Title = "Search and replace in formula Math"
		
		self.create_element('FixedText', 'label_find_symbols', {'Label' : self.msg_1})
		self.create_element('Edit', 'str_find_symbols', {})
		self.create_element('FixedText', 'label_replace_symbols', {'Label' : self.msg_2})
		self.create_element('Edit', 'str_replace_symbols',{})
		self.create_element('Button', 'btn_OK', {'Label' : 'OK', 'PushButtonType' : 1,})
		self.create_element('Button', 'btn_cancel', {'Label' : 'Cancel','PushButtonType' : 2,})
		
		
		dialog_win = self.smgr.createInstance("com.sun.star.awt.UnoControlDialog")
		dialog_win.setModel(self.dialog_win_model)
		symb_old_field = dialog_win.getControl('str_find_symbols')
		symb_new_field = dialog_win.getControl('str_replace_symbols')
		h = 25
		y = 10
		for elem in dialog_win.getControls():
			elem.setPosSize(10, y, 270, h, POSSIZE)
			y += h
		dialog_win.setPosSize(300, 300, 300, y + h, POSSIZE)
		dialog_win.setVisible(True)
		x = dialog_win.execute()
		if x:
			self.find_and_replace(symb_old_field.getText(), symb_new_field.getText())
		else:
			return False


def main():
	a = Main()
