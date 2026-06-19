"""
app.py - Flask routing only.
Maps URLs to HTML templates.
All tool logic lives in tool_logic/.
"""

from flask import Flask, render_template
from tool_logic.chain_locker import chain_locker_bp
from tool_logic.section_modulus import section_modulus_bp

app = Flask(__name__)

# Register tool blueprints.
# Add one line here per new tool.
app.register_blueprint(chain_locker_bp)
app.register_blueprint(section_modulus_bp)


# Page routes
@app.route('/')
def homepage():
    return render_template('homepage.html')


@app.route('/naval-architecture')
def basic_naval():
    return render_template('basic_naval.html')


@app.route('/structural')
def structural():
    return render_template('structural.html')


@app.route('/electrical')
def electrical():
    return render_template('electrical.html')


@app.route('/mechanical')
def mechanical():
    return render_template('mechanical.html')


# Tool pages
@app.route('/naval-architecture/chain-locker')
def chain_locker():
    return render_template('chain_locker.html')


@app.route('/structural/section-modulus')
def section_modulus():
    return render_template('section_modulus.html')


if __name__ == '__main__':
    app.run(debug=True)
