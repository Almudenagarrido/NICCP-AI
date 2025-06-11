// app/components/header.js
"use client";
import Link from "next/link";
import styles from "./header.module.css";

export default function Header() {
  return (
    <header className={styles.headerContainer}>
      <nav className={styles.navContainer}>
        <ul className={styles.navList}>
          <li className={styles.navItem}>
            <Link className={styles.navLink} href="/">Introduction</Link>
          </li>
          <li className={styles.navItem}>
            <Link className={styles.navLink} href="/execution">Execution</Link>
          </li>
          <li className={styles.navItem}>
            <Link className={styles.navLink} href="/summary-dashboard">Summary Dashboard</Link>
          </li>
          <li className={styles.navItem}>
            <Link className={styles.navLink} href="/carbon-credits">Carbon Credits</Link>
          </li>
          <li className={styles.navItem}>
            <Link className={styles.navLink} href="/techno-economic-inputs">Techno-economic Inputs</Link>
          </li>
          <li className={styles.navItem}>
            <Link className={styles.navLink} href="/cleanstep-plan">CleanStep Plan</Link>
          </li>
          <li className={styles.navItem}>
            <Link className={styles.navLink} href="/aligned-plan">Aligned-Plan</Link>
          </li>
        </ul>
      </nav>
    </header>
  );
}
