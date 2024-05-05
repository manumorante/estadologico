import Image from "next/image"
import styles from "./styles.module.scss"

export default function Juegos({ children }) {
  return (
    <div className={styles.juegos}>
      <Image src="/img/sonic.jpg" alt="Juegos Flash" width={173} height={363} />

      {children}
    </div>
  )
}
