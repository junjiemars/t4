from distutils.core import setup
import py2exe

setup(console=['t4.py'],

			name='t4',
			author='Junjie Mars',
			author_email='junjiemars@gmail.com',
			version='1.0.0',

			options={
				'py2exe':{
					'optimize': 2
					##'packages':['xml','xlwt']
					##,
					##'includes':['xlwt']
				}
			}
		 )
