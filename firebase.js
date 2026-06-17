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
  apiKey:            "PEGA_AQUI_API_KEY",
  authDomain:        "PEGA_AQUI_AUTH_DOMAIN",
  projectId:         "PEGA_AQUI_PROJECT_ID",
  storageBucket:     "PEGA_AQUI_STORAGE_BUCKET",
  messagingSenderId: "PEGA_AQUI_SENDER_ID",
  appId:             "PEGA_AQUI_APP_ID",
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
