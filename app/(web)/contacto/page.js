import Image from "next/image"
import styles from "./styles.module.scss"

export default function Contacto() {
  return (
    <div className={styles.contact}>
      <Image src="/img/latas.jpg" alt="Contacto" width={220} height={139} />

      <div className={styles.body}>
        <div className={styles.title}>Contacto</div>
        <p>
          Contacte con nosotros para <strong>contratar servicios</strong> o
          resolver alguna duda. Estaremos encantados de atenderle.
        </p>
      </div>
    </div>
  )
}
