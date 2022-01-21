import React,{useEffect,useState} from 'react'
import { getAuth, deleteUser } from "firebase/auth";
import "./styles.css"
import {useNavigate} from 'react-router-dom'
const MainPage=()=>{
    const [viewDelete,setviewDelete]=useState(false);
    const [displayName,setdisplayName]=useState("")
    const [deletionStatus,setdeletionStatus]=useState(false)
    const [errorStatus,seterrorStatus]=useState(false)
    const auth=getAuth();
    const user=auth.currentUser
    let navigate=useNavigate()
    const sleep = (milliseconds) => {
        return new Promise(resolve => setTimeout(resolve, milliseconds))
      }
    useEffect(()=>{
        let data=sessionStorage.getItem('reloadState')
        if(data==="false"){
            navigate("/")
        }
    },)
    useEffect(()=>{
        setdisplayName(sessionStorage.getItem("userName"))
    },[])
    useEffect(()=>{
        let data=sessionStorage.getItem('clearedState')
        if(data!=="true"){
            navigate("/")
        }
    },)
    useEffect(()=>{
    window.addEventListener('beforeunload', function (e) {
        sessionStorage.setItem('reloadState','false')
      });
    },[]); 
    async function handleSubmitEditForms(){
        navigate("/EditForms")
    }
    async function handleSubmitHistorical(){
        navigate("/Historical")
    }
    async function handleSubmitGenerate(){
        navigate("/Generate")
    }
    async function addUsers(){
        navigate("/Register")
    }
    const handleSubmit=(e)=>{
        e.preventDefault();
        setviewDelete(true);
    }
    const handleCancel=(e)=>{
        e.preventDefault();
        setviewDelete(false);
        setdeletionStatus(false)
        seterrorStatus(false)
    }
    const handleDelete=(e)=>{
        e.preventDefault();
        adelete()
    }
    const adelete=()=>{
        deleteUser(user).then(() => {
            setdeletionStatus(true)
            sleep(1000).then(()=>{
                navigate("/")
              })
          }).catch((error) => {
              seterrorStatus(true)
          });
    }
    const Deletion=()=>{
        if(deletionStatus===true){
            return(
                <div className="DeleteSuccess">Account has been Deleted</div>
            )
        }
        else if(errorStatus===true){
            return(
                <div className="DeleteFailure">An error has occured</div>
            )
        }
        else{
            return null
        }
    }
    const DeleteAccount=()=>{
        if(viewDelete===true){
            return(
          <section>
          <div className="Popup"></div>
          <div className="PopupText">Are you sure?</div>
          <div className="DeleteInfo">Account deletion is permanent.</div>
          <div className="DeleteInfo2">Only Administrators can create additional accounts.</div>
          <form>
            <button onClick={handleDelete}className="DeleteConfirm">Confirm</button>
            <button onClick={handleCancel}className="DeleteCancel">X</button>
          </form>
          <Deletion />
          </section>
            )
        }
        else{
            return null;
        }
    }
    const Admin=()=>{
        if(sessionStorage.getItem("userName")==='postmaster@us.af.mil'){
            return(
                <form onSubmit={addUsers}>
                <button className="MainPageOptionSelector4">Create User</button>
                </form>
            )
        }
        else{
            return(
                <form onSubmit={handleSubmit}>
                <button className="MainPageOptionSelector4">Delete Account</button>
                </form>
            )
        }
    }
    return(
        <section>
        <div>Logged in as: {displayName}</div>
        <div className="MainPageCenterBox" />
        <form onSubmit={handleSubmitEditForms}>
        <button className="MainPageOptionSelector1">Edit PS Form 3883</button>
        </form>
        <form onSubmit={handleSubmitHistorical}>
        <button className="MainPageOptionSelector2">View PS Form 3883</button>
        </form>
        <form onSubmit={handleSubmitGenerate}>
        <button className="MainPageOptionSelector3">Generate</button>
        </form>
        <Admin/>
        <DeleteAccount />
        </section>
    )
    }
export default MainPage