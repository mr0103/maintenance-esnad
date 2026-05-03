import { initializeApp } from "firebase/app";
import { getFirestore, collection, getDocs, deleteDoc, doc } from "firebase/firestore";
import fs from "fs";

const config = JSON.parse(fs.readFileSync("./firebase-applet-config.json", "utf8"));
const app = initializeApp(config);
const db = getFirestore(app, config.firestoreDatabaseId);

async function checkUsers() {
  console.log("Fetching users...");
  const snap = await getDocs(collection(db, "users"));
  snap.forEach(doc => {
    console.log("User:", doc.id, doc.data());
  });
  
  // Try to delete a test user if it exists
  const testUser = snap.docs.find(d => d.id !== "admin");
  if (testUser) {
    console.log("Attempting to delete user:", testUser.id);
    await deleteDoc(doc(db, "users", testUser.id));
    console.log("Delete command sent successfully.");
    
    // Verify deletion
    const snapAfter = await getDocs(collection(db, "users"));
    const stillExists = snapAfter.docs.some(d => d.id === testUser.id);
    console.log("User still exists after deletion:", stillExists);
  } else {
    console.log("No non-admin users found.");
  }
}

checkUsers().catch(console.error);
