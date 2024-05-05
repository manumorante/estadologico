import games from "../../../games.json"
import styles from "../styles.module.scss"
import Link from "next/link"
import Image from "next/image"

export default function Juego({ params }) {
  const { id } = params
  const game = games.find((juego) => juego.id === id)

  if (!game) {
    return <div>404</div>
  }

  return (
    <div className={styles.body}>
      <div className={styles.title}>
        <Link href="/juegos">Juegos</Link> / <strong>{game.title}</strong>
      </div>

      <div className={styles.images}>
        {Array.from({ length: game.images }).map((_, i) => (
          <img
            key={i}
            src={`/img/juegos/${id}-${i + 1}.jpg`}
            alt={game.title}
            className={styles.image}
          />
        ))}
      </div>

      <div className={styles.info}>
        <div dangerouslySetInnerHTML={{ __html: game.description }} />

        <div>
          <Link target="blank" href={`/full/${id}`} >
            <Image src="/img/jugar.gif" width={133} height={60} alt="Jugar" />
          </Link>
        </div>

        <p>
          <strong>Controles</strong>
        </p>
        <div dangerouslySetInnerHTML={{ __html: game.controls }} />
      </div>
    </div>
  )
}
