import { initializeApp } from "firebase/app";
import { getFirestore, collection, getDocs, query, where } from "firebase/firestore";
import fs from "fs";

const config = JSON.parse(fs.readFileSync("./firebase-applet-config.json", "utf8"));
const app = initializeApp(config);
const db = getFirestore(app, config.firestoreDatabaseId);

async function checkRequests() {
  console.log("Fetching hurr requests...");
  const snap = await getDocs(query(collection(db, "requests"), where("sectionId", "==", "hurr")));
  console.log(`Found ${snap.docs.length} requests.`);
  
  let malformedCount = 0;
  snap.forEach(doc => {
    const data = doc.data();
    let issues = [];
    if (!data.timestamp) issues.push("Missing timestamp");
    if (!data.answers) issues.push("Missing answers");
    if (!data.submitterName) issues.push("Missing submitterName");
    
    if (issues.length > 0) {
      console.log(`Request ${doc.id} issues:`, issues, data);
      malformedCount++;
    }
  });
  
  console.log(`Found ${malformedCount} malformed requests.`);
}

checkRequests().catch(console.error);
