"use client";
import Image from "next/image";
import styles from "./header.module.css";

export default function Header() {
  return (
    <header className={styles.headerContainer}>
      <div className={styles.headerContent}>
        <h1 className={styles.title}>
          National Integrated Clean Cooking Planning (NICCP)
        </h1>
        <div className={styles.logoGroup}>
          <Image
            src="/niccp-logo.png"
            alt="NICCP Logo"
            width={50}
            height={50}
            className={styles.logo}
          />
          <Image
            src="/se4all-logo.png"
            alt="SE4All Logo"
            width={45}
            height={45}
            className={styles.logo}
          />
          <Image
            src="/iit-logo.png"
            alt="IIT Logo"
            width={65}
            height={35}
            className={styles.logo}
          />
        </div>
      </div>
    </header>
  );
}
