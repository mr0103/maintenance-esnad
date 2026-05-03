import { initializeApp } from "firebase/app";
import { getFirestore, collection, getDocs, deleteDoc, doc } from "firebase/firestore";
import fs from "fs";

const config = JSON.parse(fs.readFileSync("./firebase-applet-config.json", "utf8"));
const app = initializeApp(config);
const db = getFirestore(app, config.firestoreDatabaseId);

async function deleteUser1() {
  console.log("Attempting to delete user '1'...");
  await deleteDoc(doc(db, "users", "1"));
  console.log("Delete command sent successfully.");
  
  const snap = await getDocs(collection(db, "users"));
  const stillExists = snap.docs.some(d => d.id === "1");
  console.log("User '1' still exists:", stillExists);
}

deleteUser1().catch(console.error);
