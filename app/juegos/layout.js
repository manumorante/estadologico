import Image from "next/image"
import Link from "next/link"
import styles from "./styles.module.scss"
import games from "./games.json"

export default function Juegos({ children }) {
  return (
    <div className={styles.juegos}>
      <Link href="/juegos">
        <Image
          src="/img/sonic.jpg"
          alt="Juegos Flash"
          width={173}
          height={363}
        />
      </Link>

      {children}

      <div className={styles.gamelist}>
        <nav>
          {games.map(({ id, title }) => (
            <Link href={`/juegos/${id}`}>{title}</Link>
          ))}
        </nav>
      </div>
    </div>
  )
}
