import React,{useState,useRef,useEffect} from 'react'
import './styles.css'
import {useNavigate} from "react-router-dom"
const Recovery=()=>{
  let navigate=useNavigate()
  const [showUsers,setshowUsers]=useState(false)
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
    const loginEmailRef=useRef(null)
    const handleSubmit=(e)=>{
      e.preventDefault();
    }
    const userView=()=>{
      setshowUsers(true)
    }
    const exitUsers=()=>{
      setshowUsers(false)
    }
    const UserList=()=>{
      if(showUsers===true){
        return(
          <section>
          <div className="DeletePopup"></div>
          <div className="DeleteText">Users</div>
          <form onSubmit={handleSubmit}>
            <button onClick={exitUsers}className="ExitUsers">X</button>
          </form>
          </section>
        )
      }
      else{
        return null
      }
    }
    
    return(
        <section>
      <nav className="navigation">DBAMS</nav>
      <div className="centerBoxLogin"></div>
      <div className="centerTextRecovery">Delete User</div>
      <form onClick={handleSubmit}>
      <input className="userEmailRecovery" type="text" placeholder="Enter Email...." ref={loginEmailRef} />
      <button className="submittyRecovery">Submit</button>
      <button onClick={userView}className="showUsers">Show Users</button>
      </form>
      <UserList />
    </section>
    )
}
export default Recovery;