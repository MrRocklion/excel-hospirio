import firebase_admin
from firebase_admin import credentials
from firebase_admin import firestore
import os
#agregamos las credenciales de firebase
cred = credentials.Certificate('credentials.json')
app_firebase = firebase_admin.initialize_app(cred,
                                             {
    'storageBucket': 'app-mantenimiento-91156.appspot.com'
})
db = firestore.client()

def ver_db():
    docs = db.collection("ingreso").stream()
    docs_mantenimientos = []
    for doc in docs:
        aux = doc.to_dict()
        try:
            if aux['departamento']['codigo']==32:
                pass
        except:
            
            documento_referencia = db.collection('ingreso').document(aux['id'])
            print('eliminamos: ->',aux['codigo'])
            #documento_referencia.delete()


ver_db()