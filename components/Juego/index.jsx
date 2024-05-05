import styles from "./styles.module.scss"
import Link from "next/link"

export default function Juego({ id, title }) {
  return (
    <div class={styles.item}>
      <Link href={"/juegos/" + id} className={styles.link}>
        <div className={styles.title}>{title}</div>
      </Link>
    </div>
  )
}
