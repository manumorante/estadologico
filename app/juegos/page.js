import styles from "./styles.module.scss"

export default function Juego() {
  return (
    <div className={styles.body}>
      <div className={styles.title}>Juegos</div>

      <p>
        Estos son algunos viejos ejemplos de juegos hechos con Flash, desde las
        primeras versiones del programa, allá por el año 1998.
      </p>

      <p>Pásalo bien :)</p>
    </div>
  )
}
