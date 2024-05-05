import "../global.css"
import styles from "./main.module.scss"
import Header from "@/components/Header"
import Footer from "@/components/Footer"

export const metadata = {
  title: "Estado Lógico by Manu Morante",
  description: "Estado Lógico by Manu Morante",
}

export default function RootLayout({ children }) {
  return (
    <html lang="es">
      <body>
        <div className={styles.content}>
          <Header />
          <main className={styles.pageBody}>{children}</main>
          <Footer />
        </div>
        {/* <div data-analytics-code='UA-34694189-6' /> */}
      </body>
    </html>
  )
}
