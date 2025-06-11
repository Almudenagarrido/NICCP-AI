"use client";
import Link from "next/link";
import { usePathname } from "next/navigation";
import styles from "./sidebar.module.css";

export default function Sidebar() {
  const pathname = usePathname();
  
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
          <li><Link href="/">Manage Case Studies</Link></li>
          <li><Link href="/execution">Execution</Link></li>
          <li><Link href="/summary-dashboard">Summary Dashboard</Link></li>
          <li><Link href="/carbon-credits">Carbon Credits</Link></li>
          <li><Link href="/techno-economic-inputs">Techno-economic Inputs</Link></li>
          <li><Link href="/cleanstep-plan">CleanStep Plan</Link></li>
          <li><Link href="/aligned-plan">Aligned Plan</Link></li>
        </ul>
      </nav>
    </aside>
  );
}
