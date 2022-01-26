import React from 'react'
import "./styles.css"
import sadge from "./images/sadge.png"

const Error=()=>{
    return(
        <section>
        <div className="Body"></div>
        <h1 className="errorMessage">GG. Page not found</h1>
        <div className="errorImage">
        <img src={sadge} alt="sadge"/>
        <p className="errorText">Use the Logout Button to return to the login screen</p>
        </div>
        </section>
    )
 }
 export default Error