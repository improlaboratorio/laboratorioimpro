import { initializeApp } from "firebase/app";
import { getFirestore, doc, getDoc, setDoc } from "firebase/firestore";

/* =========================================================================
   CONFIGURA AQUÍ TU PROYECTO DE FIREBASE (gratis) PARA QUE LA APP SE
   COMPARTA DE VERDAD ENTRE TODOS LOS QUE LA USEN, DESDE CUALQUIER MÓVIL U
   ORDENADOR, EN TIEMPO REAL.

   Pasos (unos 5 minutos, una sola vez):
   1. Ve a https://console.firebase.google.com y crea un proyecto nuevo
      (gratis, no hace falta tarjeta).
   2. Dentro del proyecto: menú lateral -> "Compilación" -> "Firestore Database"
      -> "Crear base de datos" -> elige "Modo de prueba" para empezar rápido
      (más adelante se puede restringir el acceso si hace falta).
   3. Menú lateral -> el engranaje (Configuración del proyecto) -> baja hasta
      "Tus apps" -> icono "</>" (Web) -> ponle un nombre -> "Registrar app".
   4. Te mostrará un bloque "firebaseConfig = {...}". Copia esos valores y
      pégalos aquí abajo, sustituyendo los que dicen "TU_...".
   5. Guarda este archivo y vuelve a desplegar en Netlify (o simplemente
      sube de nuevo la carpeta) para que el cambio se aplique.

   Mientras NO rellenes esto, la app sigue funcionando pero cada persona ve
   solo sus propios datos guardados en su navegador (no se comparte nada
   entre dispositivos ni personas) — verás un aviso en la parte de arriba
   de la app recordándotelo.
   ========================================================================= */
const firebaseConfig = {
  apiKey: "TU_API_KEY",
  authDomain: "TU_PROYECTO.firebaseapp.com",
  projectId: "TU_PROYECTO",
  storageBucket: "TU_PROYECTO.appspot.com",
  messagingSenderId: "TU_SENDER_ID",
  appId: "TU_APP_ID",
};

export const firebaseReady = Boolean(firebaseConfig.apiKey && firebaseConfig.apiKey !== "TU_API_KEY");

let db = null;
if (firebaseReady) {
  const app = initializeApp(firebaseConfig);
  db = getFirestore(app);
}

const COLLECTION = "asistencia_impro";

/* Misma "forma" que la API de almacenamiento del entorno de Claude
   (window.storage), para que el resto de la app no tenga que cambiar:
   storage.get(key) -> {key, value} | null
   storage.set(key, value) -> {key, value} | null                         */
export const storage = {
  async get(key) {
    if (db) {
      try {
        const snap = await getDoc(doc(db, COLLECTION, key));
        if (!snap.exists()) return null;
        return { key, value: snap.data().value };
      } catch (e) {
        console.error("Error leyendo de Firebase:", e);
        return null;
      }
    }
    try {
      const raw = localStorage.getItem("asis_local_" + key);
      return raw ? { key, value: raw } : null;
    } catch (e) {
      return null;
    }
  },
  async set(key, value) {
    if (db) {
      try {
        await setDoc(doc(db, COLLECTION, key), { value });
        return { key, value };
      } catch (e) {
        console.error("Error guardando en Firebase:", e);
        return null;
      }
    }
    try {
      localStorage.setItem("asis_local_" + key, value);
      return { key, value };
    } catch (e) {
      return null;
    }
  },
};
