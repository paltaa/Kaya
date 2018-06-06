import React from 'react'
import _ from 'underscore'


export default class BigPhoto extends React.Component {
  constructor (props) {
    super(props)
    this.state = {
      hover: false
    }
  }

  isIphonePenca () {
    return window.innerWidth < 400
  }

  render () {
    return (
      <div
        style={
          _.extend({
            width: '100%',
            margin: '50px 5px',
            backgroundImage: `url(${this.props.photo})`,
            backgroundSize: 'cover',
            backgroundPosition: 'center',
            backgroundRepeat: 'no-repeat',
            height: '390px',
            color: 'white',
            textAlign: 'center',
            fontFamily: 'Raleway'
          }, this.props.style)}
        onMouseEnter={() => this.setState({hover: true})}
        onMouseLeave={() => this.setState({hover: false})}
        onTouchTap={this.props.onTouchTap}
      >
        <div
          style={{
            height: 390,
            width: '100%',
            backgroundColor: this.state.hover ? '#ffff00' : '#000',
            opacity: this.state.hover ? '0.2' : '0.3',
            position: 'absolute'
          }}
        >
        </div>
        <div style={{fontSize: 42, marginTop: 140, width: '100%', position: 'absolute', letterSpacing: 12, fontWeight: 600}}>
          {this.props.title}
        </div>
        <div style={{fontSize: 26, marginTop: this.isIphonePenca() ? 250 : 200, width: '100%', position: 'absolute'}}>
          {this.props.subtitle}
        </div>
      </div>
    )
  }
}
