// Import the functions you need from the SDKs you need
import { initializeApp } from "https://www.gstatic.com/firebasejs/11.2.0/firebase-app.js";
import { getFirestore } from "https://www.gstatic.com/firebasejs/11.2.0/firebase-firestore.js";

// TODO: Add SDKs for Firebase products that you want to use
// https://firebase.google.com/docs/web/setup#available-libraries

// Your web app's Firebase configuration
const firebaseConfig = {
apiKey: "AIzaSyBlzFztSsCgRnq_Wp4ABEdVw2LnhCPhBnI",
authDomain: "tvtc-8ed68.firebaseapp.com",
projectId: "tvtc-8ed68",
storageBucket: "tvtc-8ed68.firebasestorage.app",
messagingSenderId: "185502232634",
appId: "1:185502232634:web:c65a5ce8f28a8506a60503"
};

// Initialize Firebase
const app = initializeApp(firebaseConfig);

// Initialize Firestore and export it
const db = getFirestore(app);
export { db };




