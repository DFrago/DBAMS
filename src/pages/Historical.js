import React,{useRef,useState,useEffect} from 'react'
import { getFirestore, doc,updateDoc} from "firebase/firestore";
import {useNavigate} from 'react-router-dom'
import "./styles.css"
import Firebase from "../firebase/firebaseIndex"

const db=getFirestore(Firebase)
const codeRef=doc(db,"Workcenter","WorkcenterCode")
async function updateDocument(workcenterCode){
  await updateDoc(codeRef,{
    workcenterCode
})
}  


const Historical=()=>{
  const [displayName,setdisplayName]=useState("")
  const workcenterScan=useRef(null)  
  const [renderCon,setRenderCon]=useState(false)
  const [workcenterValue,setWorkcenterValue]=useState("");
  let navigate=useNavigate()
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
        let data=sessionStorage.getItem('reloadState')
        if(data==="false"){
            navigate("/")
        }
    },)

    useEffect(()=>{
    window.addEventListener('beforeunload', function (e) {
        sessionStorage.setItem('reloadState','false')
      });
    },[]);
  const handleClick=()=>{
    setWorkcenterValue(workcenterScan.current.value)
  }
  
  const handleSubmit=(e)=>{
    e.preventDefault();
    authenticate({workcenterValue})
    updateDocument({workcenterValue})
  }
  
  const authenticate=({workcenterValue})=>{
      const comm=/61366/gi;const elrs=/61367/gi; const wsa=/61360/gi; const efsf=/61361/gi;const eces=/61362/gi;
      const efss=/61363/gi; const emeds=/61364/gi;const redHorse=/61365/gi;const mctBDE=/61368/gi;const esb=/61369/gi;
      const Oss=/61370/gi;const evcc=/61371/gi; const ACC=/61372/gi; const Harris=/61373/gi; const edet=/61374/gi;
      const Hurricane=/61375/gi; const ada=/61376/gi; const aafs=/61377/gi; const Longhorn=/61378/gi;const eecs=/61379/gi;
      const Ammo=/61380/gi;const econs=/61381/gi;const emxs=/61382/gi;const mdb=/61383/gi;const efgs=/61384/gi;
      const Marines=/61385/gi;const kbr=/61386/gi;const sangster=/61387/gi;const emsg=/61389/gi;const uso=/61390/gi;
      if(
        workcenterValue.match(comm)||workcenterValue.match(elrs)||workcenterValue.match(wsa)
        ||workcenterValue.match(efsf)||workcenterValue.match(eces)||workcenterValue.match(efss)
        ||workcenterValue.match(emeds)||workcenterValue.match(redHorse)||workcenterValue.match(mctBDE)
        ||workcenterValue.match(esb)||workcenterValue.match(Oss)||workcenterValue.match(evcc)
        ||workcenterValue.match(ACC)||workcenterValue.match(Harris)||workcenterValue.match(edet)
        ||workcenterValue.match(Hurricane)||workcenterValue.match(ada)||workcenterValue.match(aafs)
        ||workcenterValue.match(Longhorn)||workcenterValue.match(eecs)||workcenterValue.match(Ammo)
        ||workcenterValue.match(econs)||workcenterValue.match(emxs)||workcenterValue.match(mdb)
        ||workcenterValue.match(efgs)||workcenterValue.match(Marines)||workcenterValue.match(kbr)
        ||workcenterValue.match(sangster)||workcenterValue.match(emsg)||workcenterValue.match(uso)
        ){ 
      setRenderCon("yes")
    }
    else{
       setRenderCon("no")
    }
  }
  const RenderFinal=(renderCon)=>{
    const navigate=useNavigate();
    if(renderCon==="yes"){
      return(
        <section>
        {navigate("/PullTable")}
        </section>
      )
    }
    else if(renderCon==="no"){
      return(
        <div className="incorrectWorkCenter">Unknown Workcenter</div>
      )
    }
    else{
      return <div></div>
    }
  }
  
  return(
      <section>
      <div className="Body"></div>
      <div>Logged in as: {displayName}</div>
      <div className="EditFormsCenterBox"></div>
      <div className="EditFormsCenterText">Scan Workcenter Barcode below</div>
      <form onSubmit={handleSubmit}>
      <input className="EditFormsWorkcenter" type="text" placeholder="Workcenter...." ref={workcenterScan} />
      <button className="EditFormsSubmitty" onClick={handleClick}>Submit</button>
      </form>
      {RenderFinal(renderCon)}
      </section>
    )
}
export default Historical