#Author-Harry Keen
#Description-Generates an Opendesk dxf drawing from a Fusion360 model.

import adsk.core, adsk.fusion, adsk.cam, traceback

#import system modules
import os, sys 

#get the path of add-in
ADDIN_PATH = os.path.dirname(os.path.realpath(__file__))
#print(ADDIN_PATH)

#add the path to the searchable path collection
if not ADDIN_PATH in sys.path:
  sys.path.append(ADDIN_PATH)

from xlrd import open_workbook
#import numpy
#import dxfarc
# import operator
import math


FEATURE_DICT = {'HOLES': [], 'TOPCUTINSIDE': [], 'TOPCUTOUTSIDE': [], 'TOPPOCKETINSIDE': []}
# REVCUTINSIDE
# TOPCUTINLINE
# REVCUTOUTSIDE
# REVCUTINLINE
# TOPCUTOUTSIDE
# REVHOLES
# TOPPOCKETINSIDE
# TOPCUTINSIDE
# TOPHOLES
# TOPPOCKETOUTSIDE
# REVPOCKETOUTSIDE
# REVPOCKETINSIDE
LAYER_DICT = {}
MODEL_LAYER_DICT = {}

def run(context):
  
  try:
    global LAYER_DICT
    LAYER_DICT = import_xlsx(os.path.join(ADDIN_PATH, 'assets', 'LAYERCOLOURS - new.xlsx'))      
    
    app = adsk.core.Application.get()
    ui = app.userInterface
    
    design = app.activeProduct

    #get the root component of the active design.
    rootComp = design.rootComponent

    #print(rootComp.bRepBodies.count)

    for body in rootComp.bRepBodies:

    #find face features
    #for feature in rootComp.features:

      all_faces = body.faces
      #print(all_faces.count)
      
      faces_areas = []
      
      for face in all_faces:
        faces_areas.append(face.area)
        #print(face.area)

      #find two largest face areas - top and bottom face
      #if both are the same (part with no features)
      #back face is the lower one

      biggest_area = max(faces_areas)
      biggest_area_index = faces_areas.index(biggest_area)
      #print(biggest_area_index)

      temp_area_list = faces_areas
      temp_area_list[biggest_area_index] = 0.0
      next_biggest_area = max(temp_area_list) #max(n for n in faces_areas if n != max(faces_areas))
      next_biggest_area_index = faces_areas.index(next_biggest_area)
      #print(next_biggest_area_index)

      # print('big ' + str(biggest_area))
      # print('next ' + str(next_biggest_area))

      if ("%.4f" % biggest_area) == ("%.4f" % next_biggest_area):
        #part with no features on it - take bottom face as one with lowest z value (not great solution as stuff could be orientated in a different place.)

        biggest_area_point = all_faces[biggest_area_index].pointOnFace.asArray()
        next_biggest_area_point = all_faces[next_biggest_area_index].pointOnFace.asArray()

        # print('1 - ' + str(biggest_area))
        # print('2 - ' + str(next_biggest_area))

        if biggest_area_point[2] < next_biggest_area_point[2]:
          
          back_face_ind = biggest_area_index
          top_face_ind = next_biggest_area_index
        else:
          back_face_ind = next_biggest_area_index
          top_face_ind = biggest_area_index
      else:
        back_face_ind = biggest_area_index
        top_face_ind = next_biggest_area_index

          
      #print(str(largest_area))
      #print(str(back_face_ind))
      
      back_face = all_faces[back_face_ind]
      top_face = all_faces[top_face_ind]
      #print(str(back_face.loops.count))

      #get point for referecne on top face
      top_face_point = top_face.pointOnFace

      for loop in back_face.loops:
        if loop.isOuter:
          #outside profile of part
          #print('outer profile')

          outer_profile = get_outer_profile(loop, back = True)

          depth = 10 * get_depth(back_face, top_face_point)

          layer_key = 'TOPCUTOUTSIDE'
          #layer_name = 'TOP-HOLE-' + '%.3f' % (2*radius) + 'MM-DIAM_' + '%.3f' % depth + 'MM'
          layer_name = 'TOP-CUT-OUTSIDE-' + str(int(depth)) + 'MM'

          add_cut(layer_key, layer_name, depth, outer_profile)
        
        #inner loops
        else:
          for edge in loop.edges:
            if edge.geometry.curveType == 2:
              #THRU HOLE
              #print('thru hole')
              depth = 10 * get_depth(face, top_face_point)
              center = [i * 10 for i in edge.geometry.center.asArray()]
              diam = ("%.1f" % round(20 * edge.geometry.radius, 1))
              #print('hole ' + str(depth))

              layer_key = 'TOPHOLES'
              #layer_name = 'TOP-HOLE-' + '%.3f' % (2*radius) + 'MM-DIAM_' + '%.3f' % depth + 'MM'
              layer_name = 'TOP-HOLE-' + str(diam[0]) + 'MM-DIAM_' + str(int(depth)) + 'MM'
              layer_col = layer_colour(layer_key, depth)
              
              make_model_layer_dict(layer_name, layer_col)
              #print(layer_key, layer_name, layer_col)

              hole_dict = {'layer_name': layer_name, 'center': center, 'diameter': diam}
              FEATURE_DICT['HOLES'].append(hole_dict)

              break

            else:
              #INSIDE THRU CUT
              #print('thru inside cut')

              outer_profile = get_outer_profile(loop, back = True)

              depth = 10 * get_depth(back_face, top_face_point)
              layer_key = 'TOPCUTINSIDE'
              #layer_name = 'TOP-HOLE-' + '%.3f' % (2*radius) + 'MM-DIAM_' + '%.3f' % depth + 'MM'
              layer_name = 'TOP-CUT-INSIDE-' + str(int(depth)) + 'MM'

              add_cut(layer_key, layer_name, depth, outer_profile)

              break
              

      #top and bottom faces index
      exclude_faces = [biggest_area_index, next_biggest_area_index]
      
      for i, face in enumerate(all_faces):
        if face.geometry.surfaceType == 0: #planar surface
          if not face.geometry.normal.isEqualTo(top_face.geometry.normal):
            exclude_faces.append(i)
        else:
          exclude_faces.append(i)
      #print(exclude_faces)
      
      #iterate through all faces that arn't exluded
      for index, face in enumerate(all_faces):
        if index not in exclude_faces:
          #print(index)
          for loop in face.loops:
            if loop.isOuter:
              #out_profile = loop.edges

              for edge in loop.edges:
                if edge.geometry.curveType == 2:
                  #NON THRU HOLE
                  #find distance between faces via dot product.

                  depth = 10 * get_depth(face, top_face_point)
                  center = [i * 10 for i in edge.geometry.center.asArray()]
                  diam = ("%.1f" % round(20 * edge.geometry.radius, 1))
                  #print('hole ' + str(depth))

                  layer_key = 'TOPHOLES'
                  #layer_name = 'TOP-HOLE-' + '%.3f' % (2*radius) + 'MM-DIAM_' + '%.3f' % depth + 'MM'
                  layer_name = 'TOP-HOLE-' + str(diam[0]) + 'MM-DIAM_' + str(int(depth)) + 'MM'
                  layer_col = layer_colour(layer_key, depth)
                  
                  make_model_layer_dict(layer_name, layer_col)
                  #print(layer_key, layer_name, layer_col)

                  hole_dict = {'layer_name': layer_name, 'center': center, 'diameter': diam}
                  FEATURE_DICT['HOLES'].append(hole_dict)

                  break

                else:
                  #INSIDE NON THRU CUT
                  #print('inside cut')

                  outer_profile = get_outer_profile(loop, back = False)

                  depth = 10 * get_depth(face, top_face_point)

                  layer_key = 'TOPPOCKETINSIDE'
                  #layer_name = 'TOP-HOLE-' + '%.3f' % (2*radius) + 'MM-DIAM_' + '%.3f' % depth + 'MM'
                  layer_name = 'TOP-POCKET-INSIDE-' + str(int(depth)) + 'MM'
                  
                  add_cut(layer_key, layer_name, depth, outer_profile)

                  break

            #inner loops
            else:
              print('some internal feature') #don't need to worry about this.
    
    for keys in FEATURE_DICT.keys():
      n = 0
      for arrays in FEATURE_DICT[keys]:
        
        n += 1

      # print(keys + ' ' + str(n))
    

    #print(MODEL_LAYER_DICT)

    dxf_list = gen_dxf_list(FEATURE_DICT)
    write_dxf(dxf_list)
    #ui.messageBox('done!')

  except:
    if ui:
      print('Failed:\n{}'.format(traceback.format_exc()))


def write_dxf(dxf_list):
  file_loc = '/Users/harry/Documents/github/opendesk-on-demand/wip_docs/dxf-ouputs/test.dxf'

  with open(file_loc, "w"):
      pass    
  
  dxf = open(file_loc, 'w')

  for line in dxf_list:
    dxf.write("%s\n" % line)


def get_outer_profile(loop, back):
  outer_profile_dict = {}
  counter = 0

  # get the bulge values in the correct vertex. All curves on the back face use start vertex and all non back faces use end vertex.
  # print(loop.edges.count)
  for edge in loop.edges:
  #for i in range(0,loop.edges.count - 1):
    #edge = loop.edges[i]
    if back == True:
      p = edge.startVertex.geometry
      point = [i * 10 for i in p.asArray()]
    else:
      p = edge.endVertex.geometry
      point = [i * 10 for i in p.asArray()]
    
    #if theres a bulge make it add bulge factor
    if edge.geometry.curveType == 1:
      point.append(get_bulge(edge))
      # print(counter, edge.geometry.radius)

      #if a point has a bulge find out if it should be positive or negative by finding if the point lies on the face or outside the face edge loop. Also dependant on whether it is on the back face, sort this later.
      point.append(point_in_loop(edge.geometry.center, loop.face))

      #print(point)

    outer_profile_dict[counter] = point
    counter += 1

  #print(outer_profile_dict)

  keys = outer_profile_dict.keys()

  #apply posative or negative bulge dependent on if its on the face and if its on a back face.
  for key in keys:
    if len(outer_profile_dict[key]) == 5:
      if back == True:
        if outer_profile_dict[key][4] == True:
          outer_profile_dict[key][3] = - outer_profile_dict[key][3]
        else:
          outer_profile_dict[key][3] = outer_profile_dict[key][3]
      else:
        if outer_profile_dict[key][4] == True:
          outer_profile_dict[key][3] = outer_profile_dict[key][3]
        else:
          outer_profile_dict[key][3] = - outer_profile_dict[key][3]
  
  # for n in outer_profile:
  #   print('n - ' + loop.body.name + str(n))

  outer_profile = []

  for key in keys:
    outer_profile.append(outer_profile_dict[key])

  outer_profile.append([outer_profile[0][0], outer_profile[0][1], outer_profile[0][2]])

  return outer_profile

def point_in_loop(point, face):

  evaluator = face.evaluator

  (return_value, parameter) = evaluator.getParameterAtPoint(point)
  #(return_value, projectedPoint) = evaluator.getPointAtParameter(parameter)

  return_value = evaluator.isParameterOnFace(parameter)

  return return_value

def get_bulge(edge):
  # print('arc ' + str(edge.startVertex.geometry.asArray()) + ' ' + str(edge.endVertex.geometry.asArray()))
  # print('arc center ' + str(edge.geometry.center.asArray()))

  Ax = 10 * edge.startVertex.geometry.asArray()[0]
  Ay = 10 * edge.startVertex.geometry.asArray()[1]
  Bx = 10 * edge.endVertex.geometry.asArray()[0]
  By = 10 * edge.endVertex.geometry.asArray()[1]

  # print('Ax - ' + str(Ax))
  # print('Ay - ' + str(Ay))
  # print('Bx - ' + str(Bx))
  # print('By - ' + str(By))

  cenx = 10 * edge.geometry.center.asArray()[0]
  ceny = 10 * edge.geometry.center.asArray()[1]
  rad = 10 * edge.geometry.radius

  dist = ((By - Ay) * cenx - (Bx - Ax) * ceny + Bx * Ay - By * Ax)
  distance = math.sqrt(math.pow((By - Ay), 2) + math.pow((Bx - Ax), 2))
  center_to_line = math.sqrt(math.pow((dist / distance), 2))
  bulge = (rad - center_to_line) / (distance / 2)

  # print('radius - ' + str(rad))
  # print('dist - ' + str(dist))
  # print('distance - ' + str(distance))
  # print('center_to_line - ' + str(center_to_line))
  # print('bulge - ' + str(bulge))

  return bulge

def add_cut(layer_key, layer_name, depth, outer_profile):

  global LAYER_DICT
  layer_col = layer_colour(layer_key, depth)
                  
  make_model_layer_dict(layer_name, layer_col)
  #print(layer_key, layer_name, layer_col)

  cut_dict = {'layer_name': layer_name, 'points': outer_profile}
  FEATURE_DICT[layer_key].append(cut_dict)

def make_model_layer_dict(layer_name, layer_col):
  if layer_name not in MODEL_LAYER_DICT.keys():
      global MODEL_LAYER_DICT
      MODEL_LAYER_DICT[layer_name] = [str(layer_col[0]) + str(layer_col[1]) + str(layer_col[2]), layer_col[3]]

def layer_colour(layer_key, depth):

  global LAYER_DICT

  temp_array = LAYER_DICT[layer_key]
  layer_rgb = temp_array[int(depth) - 1]

  #print(layer_rgb)

  return layer_rgb

  #32-bit integer value. When used with True Color; a 32-bit integer representing a 24-bit color value. The high-order byte (8 bits) is 0, the low-order byte an unsigned char holding the Blue value (0-255), then the Green value, and the next-to-high order byte is the Red Value. Convering this integer value to hexadecimal yields the following bit mask: 0x00RRGGBB. For example, a true color with Red==200, Green==100 and Blue==50 is 0x00C86432, and in DXF, in decimal, 13132850

  #build a layercoulours.xlsx doc where each cell has an array of 4 elements [R, G, B, Autocad 256 colour]
  #add a 420 tag with the 'decimal' colour value to describe the real RGB colour

    #     2
    # tester
    #  70
    #      0
    #  62
    #    220
    # 420
    #  16711865

def import_xlsx(file_loc):
  #file_loc = '/Users/harry/Dropbox (OpenDesk)/06_Production/06_Software/CADLine Plugin/excel files/LAYERCOLOURS - new.xlsx'
  wb = open_workbook(file_loc)

  sheet = wb.sheet_by_index(0)

  #row = sheet.row(4)

  sheetdict = {}
  for colnum in range(1, sheet.ncols):
    col_values_list = []
    for rownum in range(1, sheet.nrows):  
      #TO DO loop through each row and append in to a []
      col_values_list.append(eval(sheet.cell_value(rownum, colnum)))
    
    #print(col_values_list)
    sheetdict[sheet.cell_value(0, colnum)] = col_values_list
  #print(sheetdict.keys())
  #print(sheetdict)

  return sheetdict

def gen_dxf_list(FEATURE_DICT):

  dxf_list = []
  dxf_list = start_section(dxf_list)

  dxf_list = start_layer(dxf_list)
  
  for layer_element in MODEL_LAYER_DICT:
    #print(layer_element)
    dxf_list = add_layer(dxf_list, layer_element, MODEL_LAYER_DICT[layer_element])
  
  dxf_list = end_layer(dxf_list)
  dxf_list = end_section(dxf_list)

  dxf_list = start_section(dxf_list)
  dxf_list = start_entities(dxf_list)

    
  for holes in FEATURE_DICT['HOLES']:
    #print(holes.count)
    #hole_dict = {'layer_name': layer_name, 'center': center, 'diameter': (2*radius)}
    dxf_list = add_hole(dxf_list, holes['diameter'], holes['center'], holes['layer_name'])

  for cuts in FEATURE_DICT['TOPCUTINSIDE']:
    #cut_inside_dict = {'layer_name': layer_name, 'points': outer_profile}
    #colour = str(6)
    start_polyline(dxf_list, cuts['layer_name'])#, colour)
    for point in cuts['points']:
      #print(point)
      dxf_list = add_vertex(dxf_list, cuts['layer_name'], point)
    end_polyline(dxf_list)

  for cuts in FEATURE_DICT['TOPCUTOUTSIDE']:

    start_polyline(dxf_list, cuts['layer_name'])#, colour)
    for point in cuts['points']:
      # print(point)
      dxf_list = add_vertex(dxf_list, cuts['layer_name'], point)
    end_polyline(dxf_list)

  for cuts in FEATURE_DICT['TOPPOCKETINSIDE']:

    start_polyline(dxf_list, cuts['layer_name'])#, colour)
    for point in cuts['points']:
      #print(point)
      dxf_list = add_vertex(dxf_list, cuts['layer_name'], point)
    end_polyline(dxf_list)

  dxf_list = end_section(dxf_list)
  dxf_list = end_dxf(dxf_list)

  return dxf_list

def get_depth(face, top_face_point):
  point = face.pointOnFace
  vector = point.vectorTo(top_face_point)
  depth = face.geometry.normal.dotProduct(vector)

  return depth

def start_section(dxf):
  dxf.append('  0')
  dxf.append('SECTION')

  return dxf

def end_section(dxf):
  dxf.append('  0')
  dxf.append('ENDSEC')

  return dxf

def end_dxf(dxf):
  dxf.append("  0")
  dxf.append("SEQEND")
  dxf.append("  0")
  dxf.append("EOF")

  return dxf

def start_layer(dxf):
  dxf.append('  2')
  dxf.append('TABLES')
  dxf.append('  0')
  dxf.append('TABLE')
  dxf.append('  2')
  dxf.append('LAYER')
  dxf.append(' 70')
  dxf.append('     4')
  dxf.append('  0')
  dxf.append('LAYER')
  dxf.append('  2')
  dxf.append('0')
  dxf.append(' 70')
  dxf.append('     0')
  dxf.append(' 62')
  dxf.append('     7')
  dxf.append('  6')
  dxf.append('CONTINUOUS')

  return dxf

def add_layer(dxf, layer_name, layer_colour):
  dxf.append('  0')
  dxf.append('LAYER')
  dxf.append('  2')
  dxf.append(layer_name)
  dxf.append(' 70')
  dxf.append('     0')
  dxf.append(' 62') #autocad colour indez
  dxf.append('   ' + str(layer_colour[1]))
  # dxf.append(' 420') #rgb value
  # print(layer_colour[0])
  # dxf.append(layer_colour[0])  # can't get working...
  dxf.append('  6')
  dxf.append('CONTINUOUS')

  return dxf


def end_layer(dxf):
  dxf.append('  0')
  dxf.append('ENDTAB')

  return dxf

def start_entities(dxf):
  dxf.append('  2')
  dxf.append('ENTITIES')

  return dxf

def start_polyline(dxf, layer):#, colour):
  dxf.append('  0')
  dxf.append('POLYLINE')
  dxf.append('8')
  dxf.append(layer)
  dxf.append(' 66')
  dxf.append('100')
  dxf.append(' 70')
  dxf.append('1')
  dxf.append(' 10')
  dxf.append('0.0')
  dxf.append(' 20')
  dxf.append('0.0')
  dxf.append(' 30')
  dxf.append('0.0')

  return dxf

def add_vertex(dxf, layer, point):
  dxf.append('  0')
  dxf.append('VERTEX')
  dxf.append('  8')
  dxf.append(layer)
  dxf.append(' 10')
  dxf.append(str(point[0]))
  dxf.append(' 20')
  dxf.append(str(point[1]))
  dxf.append(' 30')
  dxf.append(str(point[2]))
  # add bulge if its there
  if len(point) >= 4:
    dxf.append(' 42')
    dxf.append(str(point[3]))

  return dxf

def end_polyline(dxf):
  dxf.append("  0")
  dxf.append("SEQEND")

  return dxf

def add_hole(dxf, diameter, point, layer):

  x = point[0]
  y = point[1]
  z = point[2]

  dxf.append("  0")
  dxf.append("CIRCLE")
  dxf.append("  8")
  dxf.append(layer)
  dxf.append(' 10')
  dxf.append(str(x))
  dxf.append(' 20')
  dxf.append(str(y))
  dxf.append(' 30')
  dxf.append(str(z))
  dxf.append(" 40")
  dxf.append(str(diameter))

  return dxf








