#version 110
uniform sampler2D texture0;
uniform vec4 tint;
void main()
{
    vec4 c = texture2D(texture0, gl_TexCoord[0].xy);
    gl_FragColor = c*tint;
}