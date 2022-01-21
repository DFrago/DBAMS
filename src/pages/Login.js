import React,{useState,useRef,useEffect} from 'react'
import './styles.css'
import {useNavigate} from "react-router-dom"
import {getAuth,signInWithEmailAndPassword} from "firebase/auth"

export default function Login(){
  const [loginEmail,setLoginEmail]=useState("")
  const [loginPassword,setLoginPassword]=useState("")
  const loginEmailRef=useRef(null)
  const loginPasswordRef=useRef(null)
  const navigate=useNavigate();
  const [validation,setValidation]=useState(0);
  const sleep = (milliseconds) => {
    return new Promise(resolve => setTimeout(resolve, milliseconds))
  }
  useEffect(()=>{
  sessionStorage.setItem('creationState','true')
},)
useEffect(()=>{
  sessionStorage.setItem('clearedState','true')
},)
   const Validate=()=>{
    if(validation===0){
      return <div></div>
    }
    if(validation===2){
      return(
        <div className="validationApproved">User successfully authenticated</div>
      )
    }
    if(validation===3){
      return(
        <div className="validationStatus">Verifying login credentials...</div>
      )
    }
    else if(validation===1){
      return(
        <div className="validationDenied">Unknown Credentials</div>
      )
    }
  }

  const handleClick=()=>{
    setLoginEmail(loginEmailRef.current.value)
    setLoginPassword(loginPasswordRef.current.value)
  }
  
  const handleSubmit=(e)=>{
    e.preventDefault();
    logIn();
  }
  const logIn=()=>{
    const auth = getAuth();
    var errorCheck=false;
    setValidation(3);
    signInWithEmailAndPassword(auth, loginEmail, loginPassword)
      .then((userCredential) => {
      const user = userCredential.user;
      ignoreError(user)
      setValidation(2);
      sessionStorage.setItem('reloadState','true')
      sessionStorage.setItem('userName',loginEmailRef.current.value)
    })
    .catch((error) => {
      errorCheck=true;
      setValidation(1);
      const errorCode = error.code;
      const errorMessage = error.message;
      ignoreError(errorCode);
      ignoreError(errorMessage);
    });
    sleep(5000).then(()=>{if(errorCheck===false){
      setValidation(1)
      navigate("/MainPage")
    }})
  }
  const ignoreError=()=>{

  }
  return(
    <section>
      <nav className="navigation">DBAMS</nav>
      <div className="centerBoxLogin"></div>
      <div className="centerTextLogin">Login</div>
      <form onSubmit={handleSubmit}>
      <input className="userEmailLogin" type="text" placeholder="Enter Email...." ref={loginEmailRef} />
      <input className="userPasswordLogin" type="password" placeholder="Enter Password...." ref={loginPasswordRef} />
      <button className="submitty" onClick={handleClick}>Submit</button>
      </form>
      <Validate />
    </section>
  )
}
 


