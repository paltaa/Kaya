import React, { Component } from 'react';
import logo from './logo.svg';
import './App.css';
import BigPhoto from './components/bigphoto';
import TRAPENSES from './json/TRAPENSES'
import JsonTable from 'react-json-table'

class App extends Component {
  render() {
    return (
      <div className="App">
        <header className="App-header">
          <img src={logo} className="App-logo" alt="logo" />
          <h1 className="App-title"> Inventarios Kaya Unite </h1>
        </header>
        <p className="App-intro">
        <BigPhoto title='SEMINARIO' subtitle='Seminario 1307 Ñuñoa' photo='/leneria1.jpg' onTouchTap={this.redirectSeminario1307} />
<JsonTable rows={TRAPENSES} />
        </p>

      </div>
    );
  }
}

export default App;
