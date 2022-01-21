import React,{useState,useRef,useEffect} from 'react'
import './styles.css'
import {getAuth,createUserWithEmailAndPassword} from "firebase/auth"
import {useNavigate} from 'react-router-dom'

const Register=()=>{
    const [viewComplete,setviewComplete]=useState(false)
    const [displayName,setdisplayName]=useState("")
    const [userEmail,setUserEmail]=useState("");
    const [userPassword,setUserPassword]=useState("");
    const userEmailRef=useRef(null);
    const userPasswordRef=useRef(null);
    const navigate=useNavigate();
    const auth=getAuth();
    const [validation,setValidation]=useState(0);
    const sleep = (milliseconds) => {
    return new Promise(resolve => setTimeout(resolve, milliseconds))
  }
  useEffect(()=>{
    let data=sessionStorage.getItem('creationState')
    if(data==="false"){
        navigate("/")
    }
},)
  useEffect(()=>{
    let data=sessionStorage.getItem('clearedState')
    if(data!=="true"){
        navigate("/")
    }
  },)
useEffect(()=>{
window.addEventListener('beforeunload', function (e) {
    sessionStorage.setItem('creationState','false')
  });
},[]); 
useEffect(()=>{
  setdisplayName(sessionStorage.getItem("userName"))
},[])
  const Validate=()=>{
    if(validation===0){
      return <div></div>
    }
    if(validation===3){
      return(
        <div className="validationStatus">Submitting login credentials...</div>
      )
    }
    else if(validation===1){
      return(
        <div className="registrationDenied">User already exists, or you have provided an invalid address</div>
      )
    }
  }
    const handleSubmit=(e)=>{
    e.preventDefault()
    doRegister()
    }

    const handleClick=()=>{
        setUserEmail(userEmailRef.current.value)
        setUserPassword(userPasswordRef.current.value)
    }
    const doRegister=()=>{
        var errorC=false;
        setValidation(3)
        createUserWithEmailAndPassword(auth,userEmail,userPassword)
        .catch((error) => {
            errorC=true;
            setValidation(1)
            setviewComplete(false)
            const errorCode = error.code;
            const errorMessage = error.message;
            ignoreError(errorCode)
            ignoreError(errorMessage)
        });
        sleep(3000).then(()=>{
          if(errorC===false){
            setviewComplete(true)
          }
        })
    }
   
    const ignoreError=()=>{

    }
    const handleSubmit2=(e)=>{
      e.preventDefault();
      setValidation(0);
      setviewComplete(false);
    }
    const CreationCompleted=()=>{
      if(viewComplete===true){
        return(
          <section>
          <div className="PopupCreation"></div>
          <div className="PopupInfoCreation">Account has been successfully created.</div>
          <form onSubmit={handleSubmit2}>
            <button className="AckButton">OK</button>
          </form>
          </section>
        )
      }
      else{
        return (null)
      }
    }
    return(
        <section>
          <div>Logged in as: {displayName}</div>
          <div className="centerBoxLogin"></div>
          <div className="centerTextRegister">Register New User</div>
          <form onSubmit={handleSubmit}>
          <input className="userEmailRegister" type="text" placeholder="Enter Email...." ref={userEmailRef} />
          <input className="userPasswordLogin" type="password" placeholder="Enter Password...." ref={userPasswordRef}/>
          <button className="submitty" onClick={handleClick}>Submit</button>
          </form>
          <Validate />
          <CreationCompleted />
        </section>
    )
}
export default Register;

