import "../global.css"
import styles from "./styles.module.scss"
import Script from "next/script"

export const metadata = {
  title: "Estado Lógico by Manu Morante",
  description: "Estado Lógico by Manu Morante",
}

export default function FullLayout({ children }) {
  return (
    <html lang="es" className={styles.html}>
      <body className={styles.body}>
        {children}
        <Script src="https://unpkg.com/@ruffle-rs/ruffle" />
      </body>
    </html>
  )
}
