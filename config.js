// Firebase (SDK modular via CDN ESM) — inicialização e helpers globais
import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.0/firebase-app.js";
import { getAnalytics } from "https://www.gstatic.com/firebasejs/10.12.0/firebase-analytics.js";
import {
  getFirestore,
  collection,
  onSnapshot,
  addDoc,
  setDoc,
  doc,
  deleteDoc
} from "https://www.gstatic.com/firebasejs/10.12.0/firebase-firestore.js";

// Config do seu projeto (sem crases/espacos no databaseURL)
const firebaseConfig = {
  apiKey: "AIzaSyC8sclCXI8PgirqJXi65g2l0tHy288wrSc",
  authDomain: "marcacao-79921.firebaseapp.com",
  databaseURL: "https://marcacao-79921-default-rtdb.firebaseio.com",
  projectId: "marcacao-79921",
  storageBucket: "marcacao-79921.firebasestorage.app",
  messagingSenderId: "413885843715",
  appId: "1:413885843715:web:7e5a68f930b846e6598136",
  measurementId: "G-1362LZG6BP"
};

const app = initializeApp(firebaseConfig);

// Analytics em HTTPS/produção; ignore erros locais
let analytics;
try { analytics = getAnalytics(app); } catch {}

// Firestore
const db = getFirestore(app);

// Helpers globais usados pelo app.js
function readList(collectionName, callback) {
  try {
    const colRef = collection(db, collectionName);
    return onSnapshot(
      colRef,
      (snapshot) => {
        const list = snapshot.docs.map((d) => ({ id: d.id, ...d.data() }));
        try { callback(list); } catch (e) { console.error(e); }
      },
      (err) => console.error("onSnapshot error:", err)
    );
  } catch (err) {
    console.error("readList error:", err);
  }
}

async function write(collectionName, data) {
  try {
    const colRef = collection(db, collectionName);
    const res = await addDoc(colRef, data);
    return res && res.id;
  } catch (err) {
    console.error("write error:", err);
    return null;
  }
}

async function update(collectionName, id, data) {
  try {
    const docRef = doc(db, collectionName, id);
    await setDoc(docRef, data, { merge: true });
    return true;
  } catch (err) {
    console.error("update error:", err);
    return false;
  }
}

async function remove(collectionName, id) {
  try {
    const docRef = doc(db, collectionName, id);
    await deleteDoc(docRef);
    return true;
  } catch (err) {
    console.error("remove error:", err);
    return false;
  }
}

// Expõe no escopo global para app.js
window.db = db;
window.readList = readList;
window.write = write;
window.update = update;
window.remove = remove;