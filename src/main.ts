import "./ui/styles.css";
import "./ui/writer.css";
import "./ui/shell.css";
import "./ui/office365.css";
import "./ui/office365-fix.css";
import { Office } from "./core/Office";

const office = new Office(document.getElementById("app"));
office.start();
