import "@/css/global.css"
import main from "@/css/main.module.scss"
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
        <div className={main.content}>
          <Header />
          <main className={main.pageBody}>{children}</main>
          <Footer />
        </div>
        {/* <div data-analytics-code='UA-34694189-6' /> */}
      </body>
    </html>
  )
}
