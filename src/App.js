import { BrowserRouter as Router, Routes, Route, Link } from "react-router-dom";
import Login from "./pages/Login";
import ErrorPage from "./pages/errorPage"
import MainPage from "./pages/MainPage"
import EditForms from "./pages/EditForms"
import Historical from "./pages/Historical"
import Generate from "./pages/Generate"
import Register from "./pages/Register"
import Logout from "./pages/Logout"
import Creation from "./pages/Creation"
import Recovery from "./pages/Recovery"
import TableGenerator from "./pages/TableGenerator"
import PullTable from "./pages/PullTable"
import "./pages/styles.css"
function App() {
  return (
    <Router>
    <nav className="navAction">
        <Link to="/Logout">Log out</Link>
    </nav>
      <Routes>
        <Route path="/" element={<Login />} />
        <Route path="*" element={<ErrorPage />} />
        <Route path="/MainPage" element={<MainPage />}/>
        <Route path="/EditForms" element={<EditForms />} />
        <Route path="/Historical" element={<Historical />} />
        <Route path="/Generate" element={<Generate />} />
        <Route path="/Register" element={<Register />} />
        <Route path="/Login" element={<Login />} />
        <Route path="/Logout" element={<Logout />} />
        <Route path="/Creation" element={<Creation />} />
        <Route path="/Recovery" element={<Recovery />} />
        <Route path="/TableGenerator" element={<TableGenerator />} />
        <Route path="/PullTable" element={<PullTable />} />
      </Routes>
    </Router>
  );
}

export default App;