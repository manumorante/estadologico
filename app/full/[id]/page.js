import games from "../../games.json"
import styles from "./styles.module.scss"

export default function Juego({ params }) {
  const { id } = params
  const game = games.find((juego) => juego.id === id)

  if (!game) {
    return <div>404</div>
  }

  return (
    <>
      <object width="650" height="400" className={styles.swf}>
        <param name="movie" value={`/swf/${id}.swf`} />
        <embed src={`/swf/${id}.swf`} width="650" height="400"></embed>
      </object>
    </>
  )
}
