#!/usr/bin/env python
# -*- coding: utf-8 -*-
#------------------------------------------------
# author: junjiemars@gmail.com
# target: Transfer 4th
#------------------------------------------------

import sys
import os.path
from os.path import basename
from os.path import abspath
import codecs
import json
from xlrd import open_workbook
from xlutils.copy import copy
from shutil import copyfile

debug = 0
verbose = 0
make = 0
path = './'
conf = 't4.conf'

rules = [
	{ 'file': 's.xls',
		'name': '2014年1月',
		'rows': [2, -1],
		'op'	: 'dup',
		'dst': {
			'file': 't.xls',
			'dup' : 't0.xls',
			'name': '2014年1月',
			'rows': [2, -1]
		},
		'cells': [
			{ 's': 2, 'd': 0 },
			{ 's': 6, 'd': 1 },
			{ 's': -1, 'd': 4, 'v': '三类业务：省公司自主制定规范、自主运营的业务' }
		]
	},
	{ 'file': 'm.xls',
		'name': '电子渠道员工',
		'rows': [2, -1],
		'op'	: 'map',
		'dst': {
			'file': 't0.xls',
			'dup' : 't1.xls',
			'name': '2014年1月',
			'rows': [2, -1]
		},
		'cells': [
			{ 's': { 'k': 5, 'v': 6 }, 'd': { 'k': 1, 'v': 'XXX' }}
		]
	}
]

def debug_output(*msg):
	if debug: print('$', msg)

def version(argv):
	print('%s version %s' % (basename(argv[0]), '1.0.0.0'))
	sys.exit(0)

def usage(argv, xcode):
	print('usage: %s [-p path] [-c conf] '
			'[-d debug] [-v verbose] [-h help]' % basename(argv[0]))
	return (xcode)

def isvalid_dir(path):
	v = os.path.isdir(path) and os.path.exists(path)
	return (v)

def isvalid_file(path):
	v = os.path.isfile(path) and os.path.exists(path)
	return (v)

def load_rules(conf):
	c = codecs.open(conf, 'r', encoding='utf8').read()
	j = json.loads(c, encoding='utf8')
	debug_output('loaded rules:%s' % j)
	return (j)

def save_rules(conf, rules):
	s = json.dumps(rules, indent=2, sort_keys=False)
	codecs.open(conf, 'w', encoding='utf8').write(s)
	debug_output('made rules:%s' % conf)

def join(path, f):
	return (os.path.join(path, f))

def sheetnum(path, name):
	w = open_workbook(path)
	ss = w.sheets();
	for s in ss:
		if s.name == name:
			return (s.number)
	return (None) 

def srcrows(src, name, begin, end):
	b = open_workbook(src)
	s = b.sheet_by_name(name)
	if end == -1: end = s.nrows
	if begin > end: begin = end
	return ((b, s, begin, end))

def dstrows(dup, dst, name, begin, end):
	b = open_workbook(dup, formatting_info=True)
	s = b.sheet_by_name(name)
	if begin == -1: s.nrows

	wb = copy(b)
	ws = wb.get_sheet(sheetnum(dst, name))
	return (wb, ws, begin, end)

def src_cell(src, row, cell):
	if cell['s'] > -1:
		return (src.cell(row, cell['s']).value)
	else:
		return (cell['v'])

def trans(path, job):
	job['file'] = join(path, job['file'])
	if not (isvalid_file(job['file'])): 
		debug_output('job:%s is not a valid file' % job['file'])
		return

	(s_w, s_s, s_brow, s_erow) = srcrows(
		job['file'], job['name'], job['rows'][0], job['rows'][1])
	debug_output('open src-rows: %s|%s|%s|%s' % (s_w, s_s, s_brow, s_erow))

	job['dst']['file'] = join(path, job['dst']['file'])
	job['dst']['dup'] = join(path, job['dst']['dup'])
	copyfile(job['dst']['file'], job['dst']['dup'])
	debug_output('copied: %s' % job['dst']['dup'])

	(d_w, d_s, d_brow, d_erow) = dstrows(
		job['dst']['dup'], job['dst']['file'], job['dst']['name'], 
		job['dst']['rows'][0], job['dst']['rows'][1])
	debug_output('open dst-rows: %s|%s|%s|%s' % (d_w, d_s, d_brow, d_erow))

	dr = d_brow

	if 'dup' == job['op']:
		debug_output('dupping: %s' % job['dst']['dup'])
		for sr in range(s_brow, s_erow):
			for c in job['cells']:
				d_s.write(dr, c['d'], src_cell(s_s, sr, c)) 
			dr += 1
	elif 'map' == job['op']:
		debug_output('mapping: %s' % job['dst']['dup'])
		m = {}
		for sr in range(s_brow, s_erow):
			for c in job['cells']:
				k = s_s.cell(sr, c['s']['k']).value
				v = s_s.cell(sr, c['s']['v']).value
				m[k] = v

		(s_w, s_s, s_brow, s_erow) = srcrows(
			job['dst']['file'], job['dst']['name'],
			job['dst']['rows'][0], job['dst']['rows'][1])

		for sr in range(s_brow, s_erow):
			for c in job['cells']:
				sc = s_s.cell(sr, c['d']['k']).value
				mv = c['d']['v']
				if m.has_key(sc):
					mv = m[sc]
				d_s.write(dr, c['d']['k'], mv)	
			dr += 1
							
	d_w.save(job['dst']['dup'])
	debug_output('saved: %s' % job['dst']['dup'])

def main(argv):
	import getopt

	global debug, verbose
	global path, conf, rules, make

	try:
		(opts, args) = getopt.getopt(sys.argv[1:], 
			'p:c:mdvh', 
			['path=','conf=','make-rules='
			 'debug','verbose','help', 'version'])
	except getopt.GetoptError, err:
		print(str(err)+'!')
		return (usage(100))

	for (k, v) in opts:
		if k in   ('-p', '--path'): path = v
		elif k in ('-c', '--conf'): conf = v
		elif k in ('-m', '--make-rules'): make += 1
		elif k in ('-d', '--debug'): debug += 1
		elif k in ('-v', '--verbose'): verbose += 1
		elif k in ('-h', '--help'): return(usage(argv, 0))
		elif k in ('--version'): version(argv)
		else: return (200)

	debug_output(argv)	
	path = abspath(path)
	debug_output('''path:%s conf:%s''' % (path, conf))

	if not (isvalid_dir(path)):
		debug_output(''''%s' is not the valid path''' % (path))
		return (1)

	conf = join(path, conf)
	if make: 
		save_rules(conf, rules)
		return (0)

	if not (isvalid_file(conf)):
		debug_output(''''%s' is not the valid conf''' % (conf))	
		return (2)

	rules = load_rules(conf)
	[trans(path, job) for job in rules]
	
	return (0)
	
if __name__ == "__main__":
	sys.exit(main(sys.argv))
