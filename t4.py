#!/usr/bin/env python
# -*- coding: utf-8 -*-
#------------------------------------------------
# author: junjiemars@gmail.com
# target: Transfer 4th
#------------------------------------------------

import json, urllib
import sys
from os.path import basename
from xlrd import open_workbook
from xlutils.copy import copy

debug = 0
verbose = 0

s_rules = {
	'target': {
			'path': '',
			'sheet_index': 0,
			'beginrow': 2,
			'cells': [
				{'c': 0, 
					'source': {
						'id': '',
						'op': 'copy',
						'r': 2,
						'c': 2
					}
				},
				{'c': 1, 
					's': {
						'id': '',
						'op': 'map'
					}
				}
			]
		},
	
	'map': {
			'id': '',
			'path': '',
			'sheet': 0,
			'op': 'map',
			'key': {
				'r': 2,
				'c': 5
			}
	},

	'sources': [
		{'path': '',
			'sheet_index': 0,
			'id': ''
			'op': 'src',
			'
		}
	]
}

m_name_no = {}

def debug_output(*msg):
	if debug: print('$', msg)

def version(argv):
	print('%s version %s' % (basename(argv[0]), '1.0.0.0'))
	sys.exit(0)

def usage(argv, xcode):
	print('usage: %s {-s source} {-t target} {-m map} '
			'[-d debug] [-v verbose] [-h help]' % basename(argv[0]))
	return (xcode)

def read_map(f):
	n = 2
	r = 3
	cb = 5
	ce = cb + 1
	
	wb = open_workbook(f)
	s = wb.sheet_by_index(n)#wb.sheet_by_name(n)
	for row in range(r, s.nrows):
		k = s.cell(row, cb).value
		print(k, type(k))
		m_name_no[k] = str(int(s.cell(row, ce).value))

##	for k in m_name_no.keys():
##		print(k, m_name_no[k])

def write_target(s, t):
	src_ns = 2
	src_row = 3
	src_cols = [2, 5, 6]
	tar_ns = 0
	tar_row = 2

	s_wb = open_workbook(s, formatting_info=True, on_demand=True)
	s_s = s_wb.sheet_by_index(src_ns)

	t_wb = copy(open_workbook(t, formatting_info=True))
	t_s = t_wb.get_sheet(tar_ns)
	print(t_s.name)

	cnt = 0
	for row in range(src_row, src_row+3):#s_s.nrows):
		t_s.write(tar_row + cnt, 0, s_s.cell(row, 2).value)
		print(s_s.cell(row, 2).ctype)
		n_no = 'xxx'
		c = s_s.cell(row, 6).value
		if m_name_no.has_key(c): 
			n_no = m_name_no[c]
		t_s.write(tar_row + cnt, 1, n_no)
		cnt = cnt + 1

	t_wb.save(t)
		
def main(argv):
	import getopt

	global debug, verbose
	fsource = 'localhost'
	ftarget = '8080'
	fmap = None

	try:
		(opts, args) = getopt.getopt(sys.argv[1:], 
			's:t:m:dvh', 
			['source=','target=','map=',
			 'debug','verbose',
			 'help', 'version'])
	except getopt.GetoptError, err:
		print(str(err)+'!')
		return (usage(100))

	for (k, v) in opts:
		if k in   ('-s', '--source'): fsource = v
		elif k in ('-t', '--target'): ftarget = v
		elif k in ('-m', '--map'): fmap = v
		elif k in ('-d', '--debug'): debug += 1
		elif k in ('-v', '--verbose'): verbose += 1
		elif k in ('-h', '--help'): return(usage(argv, 0))
		elif k in ('--version'): version(argv)
		else: return (200)

	debug_output(argv)	
	debug_output('''source:%s target:%s map:%s'''
				 % (fsource, ftarget, fmap))

	read_map(fmap)
	write_target(fsource, ftarget)

	return (0)
	
if __name__ == "__main__":
	sys.exit(main(sys.argv))
