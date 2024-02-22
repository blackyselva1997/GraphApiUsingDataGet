import * as React from "react";
import styles from "./LoaderTriangle.module.scss";

const LoaderTriangle = () => {
  return (
    <div className={styles.loader}>
      <div className={styles.tri1}></div>
      <div className={styles.tri2}></div>
      <div className={styles.tri3}></div>
      <div className={styles.tri4}></div>
      <div className={styles.tri5}></div>
    </div>
  );
};

export default LoaderTriangle;
