import styles from "./styles.module.scss"
import Image from "next/image"
import Link from "next/link"

export default function Header() {
  return (
    <header className={styles.pageHeader}>
      <Image
        src="/img/estado-logico-logo-white.png"
        width={37}
        height={28}
        alt="Estado Lógico"
        className={styles.logo}
      />
      <h1 className={styles.h1}>
        <Link href="/" title="Ir al inicio">
          estadológico
        </Link>
      </h1>
      <h2 className={styles.h2}>Desarrollo web</h2>
    </header>
  )
}
