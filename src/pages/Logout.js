import React,{useEffect} from 'react'
import './styles.css'
import {useNavigate} from "react-router-dom"
import {getAuth,signOut} from "firebase/auth"

const sleep = (milliseconds) => {
    return new Promise(resolve => setTimeout(resolve, milliseconds))
  }
const Logout=()=>{
    const navigate=useNavigate();
    useEffect(()=>{
        let data=sessionStorage.getItem('clearedState')
        if(data!=="true"){
            navigate("/")
        }
    },)
    const auth = getAuth();
        signOut(auth)
        .then(()=> {
        sleep(2000).then(()=>{navigate("/Login")})
        })  
        .catch((error) => {
        const errorCode = error.code;
        const errorMessage = error.message;
        ignoreError(errorCode)
        ignoreError(errorMessage)
    });
    
    const ignoreError=()=>{

    }
    return(
        <section>
        <div className="Body"></div>
        <div className="logoutStatus">
           You have been signed out successfully, returning you to Login screen
        </div>
        </section>
    )
}
export default Logout;