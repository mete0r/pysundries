executable_path = 'c:/program files/iputty/plink.exe'
def _plink_args(hostinfo, options):
	if 'executable_path' in options:
		args = [options['executable_path']]
	else:
		args = [executable_path]

	tmp = str(hostinfo).split('@', 1)
	tmp.reverse()
	host = tmp[0].split(':', 1)
	user = tmp[1].split(':', 1)
	hostname = host[0]
	if len(host)>1:
		port = host[1]
	else:
		port = None
	username = user[0]
	if len(user)>1:
		password = user[1]
	else:
		password = None

	if username != '':
		args += ['-l', username]
	if password:
		args += ['-pw', password]
	if port:
		args += ['-P', port]
	if 'agent' in options:
		if options['agent']:
			args.append('-agent')
		else:
			args.append('-noagent')
	if 'batch' in options and options['batch']:
		args.append('-batch')
	if 'verbose' in options and options['verbose']:
		args.append('-v')
	args.append(hostname)
	return args

def plink(hostinfo, *remote_cmd, **options):
	args = _plink_args(hostinfo, options)
	args += remote_cmd
	import subprocess, win32process
	return subprocess.Popen(args,
		stdin=subprocess.PIPE,
		stdout=subprocess.PIPE,
		stderr=subprocess.PIPE,
		creationflags=win32process.CREATE_NO_WINDOW)

import pythoncom
class PlinkRequest:
	# by pythoncom.CreateGuid()
	_reg_clsid_ = '{C46455FF-319A-450F-8840-5BD4EF3AC9E7}'
	_reg_desc_ = 'PuTTY plink Request'
	_reg_progid_ = 'Plink.Request'
	_public_methods_ = ['request']
	_public_attrs_ = ['name']
	def __init__(self):
		self.name = 'Plink.Request'
		pass

	def request(self, host, cmd, callback):
		import threading
		callback_stream = pythoncom.CoMarshalInterThreadInterfaceInStream(pythoncom.IID_IDispatch, callback)
		threading.Thread(target=self._request, args=(host, cmd, callback_stream)).start()
	def _request(self, host, cmd, callback_stream):
		pythoncom.CoInitialize()
		cb = pythoncom.CoGetInterfaceAndReleaseStream(callback_stream, pythoncom.IID_IDispatch)
		callback = JavascriptDispatchMethod(cb, memid=0, lcid=0, wantreturn=False)

		try:
			p = plink(host, cmd, verbose=True, batch=True)
			if p.wait() == 0:
				callback('complete', p.stdout.read())
			else:
				callback('error', p.stderr.read())
		except:
			import traceback
			callback('error', traceback.format_exc())

		pythoncom.CoUninitialize()

class JavascriptDispatchMethod:
	def __init__(self, dispatch, memid, **options):
		self.dispatch = dispatch
		self.memid = memid
		self.options = options
		if not 'lcid' in options:
			self.options['lcid'] = 0
		if not 'wantreturn' in options:
			self.options['wantreturn'] = True
	def __call__(self, *args):
		params = []
		for arg in args:
			if isinstance(arg, str):
				params.append(unicode(arg))
			else:
				params.append(arg)
		return self.dispatch.Invoke(self.memid, self.options['lcid'], pythoncom.DISPATCH_METHOD, self.options['wantreturn'], *params)

def register():
	import win32com.server.register
	win32com.server.register.UseCommandLine(PlinkRequest)
def unregister():
	import win32com.server.register
	win32com.server.register.UnregisterServer(PlinkRequest._reg_clsid_, PlinkRequest._reg_progid_)
