from pyravendb.store import document_store

urls = "https://a.rdbguiando.ravendb.community/studio/index.html"

# use PFX file
#cert = {"pfx": "/path/to/cert.pfx", "password": "optional password"}

# use PEM file
# cert = ("/path/to/cert.pem")

path = 'C:/Users/Victor Magal/Downloads/Raven'

# use cert / key files
cert = (f'{path}/cert.crt', f'{path}/cert.key')

store =  document_store.DocumentStore(urls=urls, database="automacao-faturas", certificate=cert)
store.initialize() 

with store.open_session() as session:
   foo = session.load("foos/1")
   
