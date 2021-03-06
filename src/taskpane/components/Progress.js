import * as React from "react";
import PropTypes from "prop-types";
import { Spinner, SpinnerType } from "office-ui-fabric-react";

function Progress(props) {
  const { logo, message, title } = props;

  return (
    <section className="ms-welcome__progress ms-u-fadeIn500">
      <img width="90" height="90" src={logo} alt={title} title={title} />
      <h1 className="ms-fontSize-su ms-fontWeight-light ms-fontColor-neutralPrimary">{title}</h1>
      <Spinner type={SpinnerType.large} label={message} />
    </section>
  );
}

export default Progress;

Progress.propTypes = {
  logo: PropTypes.string,
  message: PropTypes.string,
  title: PropTypes.string,
};
