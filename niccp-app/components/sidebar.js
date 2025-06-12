"use client";
import Link from "next/link";
import { usePathname } from "next/navigation";
import { useState } from "react";
import styles from "./sidebar.module.css";

export default function Sidebar() {
  const pathname = usePathname();
  const isCaseStudy = pathname.startsWith("/case-studies/");
  const caseId = pathname.split("/")[2] || "";

  const [dropdownOpen, setDropdownOpen] = useState(true);

  if (pathname === "/") {
    return (
      <aside className={styles.sidebar}>
        <nav>
          <ul className={styles.navList}>
            <li>
              <Link href="/">Manage Case Studies</Link>
            </li>
          </ul>
        </nav>
      </aside>
    );
  }

  return (
    <aside className={styles.sidebar}>
      <nav>
        <ul className={styles.navList}>
          <li>
            <Link href="/">Manage Case Studies</Link>
          </li>
          <li>
            <button
              onClick={() => setDropdownOpen(!dropdownOpen)}
              style={{
                background: "none",
                border: "none",
                padding: 0,
                font: "inherit",
                cursor: "pointer",
                color: "inherit",
              }}
              aria-expanded={dropdownOpen}
            >
              Case Data {dropdownOpen ? "<" : ">"}
            </button>
            {dropdownOpen && (
              <ul style={{ paddingLeft: "1rem", marginTop: "0.5rem" }}>
                <li>
                  <Link href={`/case-studies/${caseId}/general-data`}>
                    General Data
                  </Link>
                </li>
                <li>
                  <Link href={`/case-studies/${caseId}/electricity-inputs`}>
                    Financial Electricity Inputs
                  </Link>
                </li>
                <li>
                  <Link href={`/case-studies/${caseId}/lpg-inputs`}>
                    Financial LPG Inputs
                  </Link>
                </li>
                <li>
                  <Link href={`/case-studies/${caseId}/carbon-inputs`}>
                    Carbon Finance Inputs
                  </Link>
                </li>
              </ul>
            )}
          </li>
          <li>
            <Link href={isCaseStudy ? `/case-studies/${caseId}/execution` : "/execution"}>
              Execution Inputs
            </Link>
          </li>
          <li>
            <Link href={isCaseStudy ? `/case-studies/${caseId}/summary-dashboard` : "/summary-dashboard"}>
              Summary Dashboard
            </Link>
          </li>
          <li>
            <Link href={isCaseStudy ? `/case-studies/${caseId}/carbon-credits` : "/carbon-credits"}>
              Carbon Credits
            </Link>
          </li>
          <li>
            <Link href={isCaseStudy ? `/case-studies/${caseId}/techno-economic-inputs` : "/techno-economic-inputs"}>
              Techno-economic Inputs
            </Link>
          </li>
          <li>
            <Link href={isCaseStudy ? `/case-studies/${caseId}/cleanstep-plan` : "/cleanstep-plan"}>
              CleanStep Plan
            </Link>
          </li>
          <li>
            <Link href={isCaseStudy ? `/case-studies/${caseId}/aligned-plan` : "/aligned-plan"}>
              Aligned Plan
            </Link>
          </li>
        </ul>
      </nav>
    </aside>
  );
}
