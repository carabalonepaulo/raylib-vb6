#version 110
uniform mat4 mvp;
uniform float time;
void main()
{
    vec4 p = gl_Vertex;
    p.x = p.x + sin(p.y*0.05 + time)*8.0;
    gl_Position = mvp*p;
    gl_FrontColor = gl_Color;
    gl_TexCoord[0] = gl_MultiTexCoord0;
}