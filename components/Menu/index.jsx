import styles from "./styles.module.scss"
import Link from "next/link"

export default function Menu() {
  return (
    <nav className={styles.menu}>
      <Link href="/">Inicio</Link>
      <Link href="/portafolio">Portafolio</Link>
      <Link href="/juegos">Juegos</Link>
      <Link href="/contacto">Contacto</Link>
    </nav>
  )
}
