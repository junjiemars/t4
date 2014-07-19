#!/usr/bin/env python
# -*- coding: utf-8 -*-
#------------------------------------------------
# author: junjiemars@gmail.com
# target: Transfer 4th
#------------------------------------------------

import sys
import os.path
from os.path import basename
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
	{ 'path': 's.xls',
		'name': '2014年1月',
		'rows': [2, 4],
		'op'	: 'dup',
		'dst': {
			'path': 't.xls',
			'dup' : 't0.xls',
			'name': '2014年1月',
			'rows': [2, 12]
		},
		'cells': [
			{ 's': 2, 'd': 0 },
			{ 's': 6, 'd': 1 }
		]
	},
	{ 'path': 'm.xls',
		'name': '电子渠道',
		'rows': [1, 82],
		'op'	: 'map',
		'dst': {
			'path': 't0.xls',
			'dup' : 't1.xls',
			'name': '2014年1月',
			'rows': [2, 12]
		},
		'cells': [
			{ 's': { 'k': 3, 'v': 5 }, 'd': { 'k': 1, 'v': 'xxx' }}
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

def run(path, rules):
	[trans(path, job) for job in rules]

def trans(path, job):
	job['path'] = join(path, job['path'])
	if not (isvalid_file(job['path'])): 
		debug_output('job:%s is not the valid file' % job['path'])
		return

	s_w = open_workbook(job['path'])
	s_s = s_w.sheet_by_name(job['name'])

	s_brow = job['rows'][0]
	s_erow = job['rows'][1]

	job['dst']['path'] = join(path, job['dst']['path'])
	job['dst']['dup'] = join(path, job['dst']['dup'])

	t_sn = sheetnum(job['dst']['path'], job['dst']['name'])
	if None == t_sn:
		debug_output('invalid dst sheet-name:%s' % job['dst']['name'])
		return

	copyfile(job['dst']['path'], job['dst']['dup'])
	d_w = copy(open_workbook(job['dst']['dup'], formatting_info=True))

	d_brow = job['dst']['rows'][0]
	d_erow = job['dst']['rows'][1]
	d_s = d_w.get_sheet(t_sn)
	dr = d_brow

	if 'dup' == job['op']:
		for sr in range(s_brow, s_erow):
			for c in job['cells']:
				sc = s_s.cell(sr, c['s']).value
				d_s.write(dr, c['d'], sc) 
			dr += 1
	elif 'map' == job['op']:
		m = {}
		for sr in range(s_brow, s_erow):
			for c in job['cells']:
				k = s_s.cell(sr, c['s']['k']).value
				v = s_s.cell(sr, c['s']['v']).value
				m[k] = v

		s_w = open_workbook(job['dst']['path'])
		s_s = s_w.sheet_by_name(job['dst']['name'])
		s_brow = job['dst']['rows'][0]
		s_erow = job['dst']['rows'][1]

		for sr in range(s_brow, s_erow):
			for c in job['cells']:
				sc = s_s.cell(sr, c['d']['k']).value
				mv = c['d']['v']
				if m.has_key(sc):
					mv = m[sc]
				print('##', sc, mv)
				d_s.write(dr, c['d']['k'], mv)	
			dr += 1
							
	d_w.save(job['dst']['dup'])

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
	run(path, rules)
	
##	read_map(fmap)
##	write_target(fsource, ftarget)

	return (0)
	
if __name__ == "__main__":
	sys.exit(main(sys.argv))
