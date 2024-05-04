import styles from "../css/footer.module.scss"
import Link from "next/link"
export default function Footer() {
  const currentYear = new Date().getFullYear()
  return (
    <footer class={styles.footer}>
      <nav>
        <ul>
          <li>
            <Link href="portafolio.html" title="Galería de trabajos realizados">
              Portafolio
            </Link>
          </li>
          <li>
            <Link href="juegos.html" title="Algunos juegos clásicos en Flash">
              Juegos
            </Link>
          </li>
          <li>
            <Link href="contacto.html" title="Eh! contacta conmigo! ;-)">
              Contacto
            </Link>
          </li>
        </ul>
      </nav>

      <small>&copy; 2001-{currentYear} Manu Morante</small>
    </footer>
  )
}
