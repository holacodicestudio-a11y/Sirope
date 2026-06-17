// ════════════════════════════════════════════════
//  CONEXIÓN A FIREBASE (la nube de Sirope)
//
//  👉 PEGA AQUÍ los datos de TU proyecto de Firebase.
//     Los obtienes en la consola de Firebase (ver la guía).
//     Reemplaza los valores "PEGA_AQUI_..." por los tuyos.
// ════════════════════════════════════════════════
import { initializeApp } from "firebase/app";
import {
  initializeFirestore,
  persistentLocalCache,
} from "firebase/firestore";
import { getAuth } from "firebase/auth";

const firebaseConfig = {
  apiKey:            "AIzaSyCUeYUE6UU8xeSTFhP3UVV9PajNOBWclTk",
  authDomain:        "sirope-f3134.firebaseapp.com",
  projectId:         "sirope-f3134",
  storageBucket:     "sirope-f3134.firebasestorage.app",
  messagingSenderId: "668210143755",
  appId:             "1:668210143755:web:860b5a71cdc487bfdc9d0e",
};

// ¿Ya está configurado? (para mostrar aviso amable si aún no)
export const FIREBASE_LISTO = !firebaseConfig.apiKey.startsWith("PEGA_AQUI");

const app = initializeApp(firebaseConfig);

// Firestore con caché offline: la app funciona sin internet
// y sincroniza sola cuando regresa la conexión.
export const db = initializeFirestore(app, {
  localCache: persistentLocalCache(),
});

export const auth = getAuth(app);
