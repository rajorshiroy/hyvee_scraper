from hyvee import Fiddler, SazParser

fiddler = Fiddler()
# fiddler.clean_fiddler_session()
fiddler.unpack_saz()


sp = SazParser()
sp.get_requests()
