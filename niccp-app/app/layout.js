// layout.js
import "./globals.css";

import Header from "/components/header";
import Sidebar from "/components/sidebar";
import Footer from "/components/footer";

export const metadata = {
  title: "NICCP-App",
  description: "National Integrates Clean Cooking Planning Software Tool",
};

export default function RootLayout({ children }) {
  return (
    <html lang="en">
      <body>
        <Header />
        <div className="layout">
          <aside> 
            <Sidebar/>
          </aside>
          <main>{children}</main>
        </div>
        <Footer />
      </body>
    </html>
  );
}
