import firebase from 'firebase/compat/app'
import 'firebase/compat/firestore'

const firebaseConfig = {
    authDomain: "dbamsstorage.firebaseapp.com",
    projectId: "dbamsstorage",
    storageBucket: "dbamsstorage.appspot.com",
    messagingSenderId: "5959517180",
    appId: "1:5959517180:web:46f58490ca88e8ad9a2879"
  };

  //Initialize Firebase
  firebase.initializeApp(firebaseConfig)
  const Firebase=firebase.initializeApp(firebaseConfig);
  export default Firebase;
