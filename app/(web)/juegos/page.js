import Link from "next/link"
import styles from "./styles.module.scss"
import games from "../../games.json"

export default function Juego() {
  return (
    <div className={styles.body}>
      <div className={styles.title}>Juegos</div>

      <p>
        Estos son algunos viejos ejemplos de juegos hechos con Flash, desde las
        primeras versiones del programa, allá por el año 1998.
      </p>

      <p>Pásalo bien :)</p>

      <div className={styles.images}>
        {games.map(({ id, title }) => (
          <Link href={`/juegos/${id}`} key={id}>
            <img
              src={`/img/juegos/${id}-1.jpg`}
              width={150}
              height="auto"
              className={styles.image}
              alt={title}
            />
            <div>{title}</div>
          </Link>
        ))}
      </div>
    </div>
  )
}
