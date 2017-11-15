README
======


``template.pptx`` is an empty powerpoint file with the slide masters set to the layout for the 2017-RSV. It has a white background.


``icons.pptx`` contains a few pages with graphic equivalents of each character from the et-line_ and font-awesome_ fonts. These graphics are not dependent on the fonts themselves, so powerpoint files with these graphics in them are portable to other systems where the fonts may not be installed.

The icons are all in vector graphic format so will scale without pixelation artifacts.


Tech Info
---------

The ``icons.pptx`` powerpoint is geneated by the python script ``power_icons.py``, which basically adds pages with matrices of ``.emf`` graphics, each graphic represents a single character from the source fonts.


The ``.emf``  files are EMF_ format files which is a vector format supported by Powerpoint (it does not directly support the open standard ``.svg`` file format).

SVG files are converted to EMF using the java tool svg2vector_ which in turn uses Inkscape_ to do the conversion grunt work. The java ``.jar`` file is wrapped with a script ``s2v``::

  #!/bin/bash
  INKSCAPE=/usr/local/bin/inkscape
  JAR=${HOME}/bin/svg2vector-2.0.0-jar-with-dependencies.jar
  java -jar ${JAR} s2v-is -x ${INKSCAPE} $@

Converting the folder of SVG to EMF was done with::

  for fn in svg/*.svg; do echo $fn; s2v -f ${fn} -d emf/ -t emf; done


SVG files are provided in the et-line_ distribution.

SVG files for font-awesome_ were generated using the Font-Awesome-SVG-PNG_ tool.

.. _EMF: https://en.wikipedia.org/wiki/Windows_Metafile
.. _et-line: https://github.com/pprince/etlinefont-bower
.. _font-awesome: http://fontawesome.io/
.. _Font-Awesome-SVG-PNG: https://github.com/encharm/Font-Awesome-SVG-PNG
.. _Inkscape: https://inkscape.org/en/
.. _svg2vector: https://github.com/vdmeer/svg2vector

