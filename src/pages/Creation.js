import React,{useEffect} from 'react'
import './styles.css'
import {useNavigate} from "react-router-dom"


const sleep = (milliseconds) => {
    return new Promise(resolve => setTimeout(resolve, milliseconds))
  }
  
const Creation=()=>{
    useEffect(()=>{
        let data=sessionStorage.getItem('clearedState')
        if(data!=="true"){
            navigate("/")
        }
    },)
    const navigate=useNavigate();
    sleep(2000).then(()=>{navigate("/Login")})
    return(
        <section>
        <div className="Body"></div>
        <div className="logoutStatus">Account successfully created, you will be returned to the Login page.</div>
        </section>
    )
}
export default Creation;