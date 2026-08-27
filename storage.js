import { initializeApp } from "firebase/app";
import { getDatabase, ref, get, set } from "firebase/database";

const firebaseConfig = {
  apiKey: "AIzaSyDv6IlqmUjlR7-6OghC1lFyQOSks6XLeTM",
  authDomain: "laboratorio-impro.firebaseapp.com",
  databaseURL: "https://laboratorio-impro-default-rtdb.europe-west1.firebasedatabase.app",
  projectId: "laboratorio-impro",
  storageBucket: "laboratorio-impro.firebasestorage.app",
  messagingSenderId: "361006442640",
  appId: "1:361006442640:web:980383f7fdc37d77309015",
};

export const firebaseReady = Boolean(firebaseConfig.apiKey && firebaseConfig.apiKey !== "TU_API_KEY");

let db = null;
if (firebaseReady) {
  const app = initializeApp(firebaseConfig);
  db = getDatabase(app);
}

const ROOT = "asistencia_impro";

export const storage = {
  async get(key) {
    if (db) {
      try {
        const snap = await get(ref(db, `${ROOT}/${key}`));
        if (!snap.exists()) return null;
        return { key, value: snap.val() };
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
        await set(ref(db, `${ROOT}/${key}`), value);
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
