Attribute VB_Name = "Universe"
'**************************************************************
'
' THIS WORK, INCLUDING THE SOURCE CODE, DOCUMENTATION
' AND RELATED MEDIA AND DATA, IS PLACED INTO THE PUBLIC DOMAIN.
'
' THE ORIGINAL AUTHOR IS SILVER FOX.
'
' THIS SOFTWARE IS PROVIDED AS-IS WITHOUT WARRANTY
' OF ANY KIND, NOT EVEN THE IMPLIED WARRANTY OF
' MERCHANTABILITY. THE AUTHOR OF THIS SOFTWARE,
' ASSUMES _NO_ RESPONSIBILITY FOR ANY CONSEQUENCE
' RESULTING FROM THE USE, MODIFICATION, OR
' REDISTRIBUTION OF THIS SOFTWARE.
'
'**************************************************************
'
' This file was downloaded from The Game Programming Wiki.
' Come and visit us at http://gpwiki.org
'
'**************************************************************

Option Explicit

Global gblnUniverseLoaded As Boolean    'Is the universe loaded?
Global gstrUniverse As String           'Name of the universe file to load
Global Const DEFAULT_UNIVERSE = "universe.uni"

'Sprite info
Type SPRITETYPE
    strResName As String
    bytFrameNum As Byte
    bytFrameAmt As Byte
    bytAnimNum As Byte
    bytAnimAmt As Byte
    lngAnimRate As Long         'Number of MS between anims
    lngAnimLast As Long         'Number of MS since last anim
    lngSpriteObject As Long     'Holds the spriteobject reference number
    intWidth As Integer
    intHeight As Integer
    blnLoaded As Boolean
End Type
Global Const FRAME_NUM = 39

'Player AI vars
Global Const ALLSTOP_SPEED = 0.02
Global Const AUTOPILOT_DIST = 10        'Number of pixels to stop within
Global Const AUTOPILOT_FTL_DIST As Double = 0.5 'Number of LY to cause a stop (and dist fudge)
Global Const AUTOPILOT_FTL_SETDIST = 1          'Number of AU to stop within

'Loading/unloading tolerances
Global Const LOAD_DISTANCE = 20000
Global Const UNLOAD_DISTANCE = 25000

'Races
Global Const RACE_NUM = 12
Global Const RACE_TERRAN = 0
Global Const RACE_KALE = 1
Global Const RACE_PRAEMALI = 2
Global Const RACE_HANTAKAS = 3
Global Const RACE_ALTAIRIAN = 4
Global Const RACE_GRAME = 5
Global Const RACE_VEGAN = 6
Global Const RACE_ULWAR = 7
Global Const RACE_TULONI = 8
Global Const RACE_SICARIUS = 9
Global Const RACE_INDEPENDENT = 10
Global Const RACE_PLAYER = 11
Global Const RACE_PLANET = 12
Type RACETYPE
    strName As String                   'What's the name of this race?
    blnEncountered As Boolean           'Has this race been encountered by the player yet?
    intRelations(RACE_NUM) As Integer   'How highly does this race view the others?
    lngGenerationTime As Long           'How many milliseconds between ship creations? (Generate at starbase.. IF there is one!)
    lngCurrentGenerationTime As Long    'How many milliseconds since last ship creation?
    intMaxShips As Integer              'How many ships can this race have maximally?
End Type
Global gudtRace(RACE_NUM) As RACETYPE

'Relations
Global Const RELATIONS_GOOD = 100       'intRelations value that is considered "friendly"
Global Const RELATIONS_BAD = -100       'intRelations value that is considered "enemy"

'Weapons
Global Const LASER_NUM = 5
Global Const CANNON_NUM = 5
Global Const MISSILE_NUM = 5
Global glngLaserNum As Long
Global glngCannonNum As Long
Global glngMissileNum As Long

'Ship parts
Global Const SHIELD_NUM = 10
Global Const HULL_NUM = 50
Global Const GENERATOR_NUM = 10
Global Const ENGINE_NUM = 10
Global Const ARMOUR_NUM = 10
Global Const SCANNER_NUM = 10
Global glngShieldNum As Long
Global glngHullNum As Long
Global glngGeneratorNum As Long
Global glngEngineNum As Long
Global glngArmourNum As Long
Global glngScannerNum As Long

'Normalizations
Global Const NORMALIZE_DISTANCE_AU = 10000  '1 AU = _____ pixels
Global Const NORMALIZE_DISTANCE_LY = 64000  '1 LY = _____ AU
Global Const NORMALIZE_SPEED = 0.65         '.99c = _____ speed

'FTLD and ARCD constants
Global Const FTLD_CONSUMPTION = 115      'Energy consumed on activation
Global Const ARCD_CONSUMPTION = 200
Global Const FTLD_SPEED = 10000000 * NORMALIZE_SPEED     'Speed of travel
Global Const ARCD_SPEED = 50000000 * NORMALIZE_SPEED
Global Const MIN_LIGHT_SPEED = 0.5 * NORMALIZE_SPEED     'Minimum requisite speed for jump to light

'Jammer
Private Type JAMMERTYPE
    sngConsumption As Single
    dblRange As Double
End Type
Global gudtJammer As JAMMERTYPE

'Engines
Type ENGINETYPE
    strName As String           'Name of the engine
    sngThrust As Single         'Thrust/unit of energy for all the engine types
    sngConsumption As Single    'Units of energy consumed per ms
    lngMaxEnergy As Long        'Maximal energy value
    lngSound As Long            'The sound index associated with this engine
End Type
Global gudtEngine(ENGINE_NUM) As ENGINETYPE
Global gstrEngineSound(ENGINE_NUM) As String

'Generators
Type GENERATORTYPE
    strName As String           'Name of the generator
    sngConsumption As Single    'Units of fuel consumed per ms
    sngOutPut As Single         'Units of energy created per ms
    lngMaxEnergy As Long        'Maximal energy value
    lngMaxBattery As Long       'Maximal battery value
End Type
Global gudtGenerator(GENERATOR_NUM) As GENERATORTYPE

'Shields
Type SHIELDTYPE
    strName As String           'Name of the shield
    sngConsumption As Single    'Units of energy consumed per ms
    sngAbsorbtion As Single     'Units of energy consumed per unit of damage
    lngMaxEnergy As Long        'Maximal energy value
End Type
Global gudtShield(SHIELD_NUM) As SHIELDTYPE

'Lasers
Type LASERTYPE
    strName As String           'Name of the laser
    sngConsumption As Single    'Units of energy consumed per ms
    sngFireConsumption As Single    'Units of energy consumed per MS during fire
    sngConcussiveDamage As Single   'Units of concussive damage per ms contact
    sngRadiationDamage As Single    'Units of radiation damage per ms contact
    lngRange As Long            'Max range of the laser
    lngMaxEnergy As Long        'Maximal energy value
    lngColour As Long           'Colour of the beam
    lngSound As Long            'Sound index associated with this laser
End Type
Global gudtLaser(LASER_NUM) As LASERTYPE
Global gstrLaserSound(LASER_NUM) As String
Global Const LASER_WIDTH = 2

'Cannons
Type CANNONTYPE
    strName As String
    sngConsumption As Single
    sngInstantaneousConsumption As Single
    sngConcussiveDamage As Single
    sngRadiationDamage As Single
    sngSpeed As Single          'Speed of projectile
    lngDuration As Long         'Lifetime of projectile in ms
    lngMaxEnergy As Long
    lngFireRate As Long         'Number of MS between shots
    lngSprite As Long           'Reference to sprite
    lngSound As Long            'Sound index associated with this cannon
    lngCannonType As Long       'What kind of cannon is it?  Single?  Double?  Etc.
End Type
Global gudtCannon(CANNON_NUM) As CANNONTYPE
Global gstrCannonSound(CANNON_NUM) As String
Global gstrCannonSprite(CANNON_NUM) As String

'Cannon Type Constants
Global Const CANNON_FIXED = 1
Global Const CANNON_TURRET = 2
Global Const CANNON_DOUBLE = 4
Global Const CANNON_TRIPLE = 8
Global Const CANNON_QUAD = 16
Global Const CANNON_OCTA = 32
'Distance between bullets for double/triple/quad/octa turret
Global Const CANNON_SPREAD = 4

'Missiles
Type MISSILETYPE
    strName As String
    lngSound As Long
    sngConcussiveDamage As Single
    sngRadiationDamage As Single
    sngRotationRate As Single
    lngFireRate As Long
    sngMaxSpeed As Single       'Max speed of the missile
    sngThrust As Single         'Thrust/unit of energy for all the engine types
    lngDuration As Long         'Lifetime of projectile in ms
    sngTargetBias As Single     'AI stuff
    sngSeekDist As Single
    udtSprite As SPRITETYPE
End Type
Global gudtMissile(MISSILE_NUM) As MISSILETYPE
Global gstrMissileSound(MISSILE_NUM) As String

'Armour
Type ARMOURTYPE
    strName As String
    lngMaxArmour As Long
End Type
Global gudtArmour(ARMOUR_NUM) As ARMOURTYPE

'Scanners
Type SCANNERTPYE
    strName As String
    dblMaxRange As Double
End Type
Global gudtScanner(SCANNER_NUM) As SCANNERTPYE
Global Const MIN_SCANNER_RANGE = 1825
Global Const RADAR_WIDTH = 120
Global Const RADAR_STEP = 0.002                  '% Radar change per ms

'Hulls
Type HULLTYPE
    strName As String
    lngMass As Long             'Mass of this hull type
    lngMaxCargo As Long         'Tons of cargo space
    lngMaxCrew As Long          'Maximal crew compliment
    lngMaxMissile As Long       'Number of missiles hull can carry
    lngMaxMines As Long         'Number of mines hull can carry
    lngMaxFuel As Long          'Max fuel capacity
    sngMaxSpeed As Single       'Max speed of hull (Relativistic effects put stress on hull.. each has different ability to cope)
    sngRotationRate As Single   'Strength of rotational thrusters
    'What component levels can this hull have?
    'Treat these as bitflags (0 = cannot have, 1 = can have lowest level, 3 = can have second and first levels, etc)
    lngLaser As Long
    lngCannon As Long
    lngMissile As Long      '(Exclusive.  If lngMissle = 1, then ONLY type 1 missiles can be used)
    lngShield As Long
    lngGenerator As Long
    lngEngine As Long
    lngArmour As Long
    'Accessories?
    blnCommJammer As Boolean
    blnMines As Boolean
    blnFTLD As Boolean
    blnARCD As Boolean
    udtSprite As SPRITETYPE
End Type
Global gudtHull(HULL_NUM) As HULLTYPE

'UDTs
Type PHYSICSTYPE
    dblX As Double
    dblY As Double
    sngFacing As Single
    sngHeading As Single
    sngSpeed As Single
    lngMass As Long
    blnThrusting As Boolean
    blnReverseThrusting As Boolean
    blnTurningLeft As Boolean
    blnTurningRight As Boolean
End Type
Type INFOTYPE
    dblDistance As Double       'Distance from the player
    bytRace As Byte
    strName As String
    blnCanMove As Boolean
    blnCarrier As Boolean       'Is this ship a carrier?
    blnFighter As Boolean       'Is this ship a fighter?
    blnStarBase As Boolean      'Is this a starbase?
    blnPlanet As Boolean        'Is this a planet?
    blnStar As Boolean          'Is this a star?
    blnShieldUp As Boolean      'Is the shield graphic displayed?
    lngShieldDown As Long       'At what gametime should the shields come down?
End Type
Type SYSTEMSTYPE
    sngRotationRate As Single   'Radians per MS
    sngFuel As Single
    lngArmour As Long
    lngCrew As Long
    sngEnergy As Single         'Current energy store
    sngGeneratorEnergy As Single    'Slider values of the various systems
    sngShieldEnergy As Single
    sngEngineEnergy As Single
    sngWeaponEnergy As Single   'Max weapon energy is the sum of lasers + cannons maximums
    bytArmour As Byte
    bytShield As Byte
    bytEngine As Byte
    bytCannon As Byte
    lngCannonLastFire As Long
    bytLaser As Byte
    bytMissile As Byte
    lngMissileLastFire As Long
    bytHull As Byte
    bytGenerator As Byte
    bytScanner As Byte
    intMissileNum As Integer
    intMineNum As Integer
    blnJammer As Boolean        'Communications jammer (stops TRANSMISSION, not reception!)
    blnJammerActive As Boolean
    blnFTLD As Boolean          'Faster than light drive?
    blnFTLDActive As Boolean
    blnARCD As Boolean          'Arc drive?
    blnARCDActive As Boolean
End Type
Type CARRIERTYPE
    intMaxFighters As Integer   'How many can it carry?
    intFighters As Integer      'How many is it currently carrying?
End Type
Type FIGHTERTYPE
    lngFighterOwner As Long     'Which carrier owns this fighter?
End Type
Global Const COMMODITY_ORE = 0
'
' ..more..
'
Global Const COMMODITY_NUM = 5  'Number of commodities (zero based)
Type STARBASETYPE
    lngCommodityPrice(COMMODITY_NUM) As Long    'Price of the commodities
    lngLaserPrice(LASER_NUM) As Long            'Price of lasers
    lngCannonPrice(CANNON_NUM) As Long
    lngMissilePrice(MISSILE_NUM) As Long
    lngHullPrice(HULL_NUM) As Long
    lngShieldPrice(SHIELD_NUM) As Long
    lngGeneratorPrice(GENERATOR_NUM) As Long
    lngEnginePrice(ENGINE_NUM) As Long
    lngArmourPrice(ARMOUR_NUM) As Long
    lngCommJammerPrice As Long
    lngMinesPrice As Long
    lngFTLDPrice As Long
    lngARCDPrice As Long
End Type
Global Const AI_NONE = 0
Global Const AI_FLEE = 1
Global Const AI_ATTACK = 2
Global Const AI_PATROL = 3
Global Const AI_TRADE = 4
Global Const AI_SEEK = 5
Global Const AI_ALLSTOP = 6
Global Const AI_AUTOPILOT = 7
Global Const TARGET_PLAYER = -1
Global Const TARGET_NONE = -2
Global Const TARGET_COORDS = -3
Global Const LASER_FIRE_MIN = 0.5       'Percent energy at which to cease laser fire
Global Const LASER_FIRE_RECHARGE = 0.75 'Percent energy before AI will fire lasers
Global Const MISSILE_FIRE_DIST = 2000   'Distance within which AI will fire a missile

Type AITYPE
    lngTarget As Long           'Set to -1 if player is target, -2 if no target
    dblX As Double              'X destination coord
    dblY As Double              'Y destination coord
    bytAction As Byte           'Flee, attack, patrol, or trade?
    'SNGTOLERANCE DEPRECATED
    sngTolerance As Single      'Radius from dblX and dblY within which object can act
    'SNGTOLERANCE DEPRECATED
    sngTargetBias As Single
    sngSeekDist As Single
    sngMinDist As Single
    sngCannonDist As Single     'Distance needed for cannon fire
    sngAimTolerance As Single   'Angle (in radians) on either side of direct line to target in which cannon fire should take place
    lngLengthTargetLock As Long 'How long do we keep the current target?
    lngNewTargetTime As Long    'When do we acquire a new target?
    blnThrusting As Boolean     'Were we thrusting? (For sound purposes)
    lngThrustSound As Long      'Reference for the currently playing thrust sound
    blnLaserFire As Boolean     'Were we firing lasers? (For sound purposes)
    lngLaserSound As Long       'Reference for the currently playing laser sound
End Type
Type CARGOTYPE
    lngCredits As Long          'Credits player has
    lngSalvage As Long          'Tons of salvage
    lngNumCargo As Long         'How many different types of cargo do we have?
    bytCommodity() As Byte      'What commodity do we have in the cargo holds?
    lngAmount() As Long         '..and how many tons of it?
End Type
Type CONTROLTYPE                'Slider values
    bytGenerator As Byte
    bytEngine As Byte
    bytShield As Byte
    bytWeapons As Byte
End Type

'Objects
Type OBJECTTYPE
    blnExists As Boolean
    udtPhysics As PHYSICSTYPE
    udtInfo As INFOTYPE
    udtSprite As SPRITETYPE
    udtSystems As SYSTEMSTYPE
    udtCarrier As CARRIERTYPE
    udtFighter As FIGHTERTYPE
    udtStarbase As STARBASETYPE
    udtAI As AITYPE
End Type
Global gudtObject() As OBJECTTYPE

'Player
Type PLAYERTYPE
    udtPhysics As PHYSICSTYPE
    udtInfo As INFOTYPE
    udtSprite As SPRITETYPE
    udtSystems As SYSTEMSTYPE
    udtCargo As CARGOTYPE
    udtCarrier As CARRIERTYPE
    udtControl As CONTROLTYPE
    udtAI As AITYPE
    dblCurrentRange As Double
    lngRadarObject As Long
End Type
Global gudtPlayer As PLAYERTYPE

'Weapons
Global Const BULLET_WIDTH = 4
Global Const BULLET_HEIGHT = 4
Type BULLETTYPE
    dblX As Double
    dblY As Double
    sngSpeed As Single
    sngDirection As Single
    bytCannon As Byte       'Type of cannon this bullet is from
    lngOwner As Long        'Who shot this? -1 = player
    lngCreated As Long      'Time at which the bullet was spawned
End Type
Global glngNumBullets As Long
Global gudtBullet() As BULLETTYPE

'Missiles
Type LIVEMISSILETPYE
    dblX As Double
    dblY As Double
    sngSpeed As Single
    sngDirection As Single
    sngFacing As Single
    bytMissile As Byte
    lngOwner As Long
    lngTarget As Long
    lngCreated As Long
    lngSmokeTime As Long
    udtSprite As SPRITETYPE
    dblXPrev(4) As Double
    dblYPrev(4) As Double
End Type
Global glngNumLiveMissiles As Long
Global gudtLiveMissile() As LIVEMISSILETPYE
Global Const SMOKE_TRAIL_DELAY = 25

'Radar
Private Type RADARTYPE
    lngObject As Long
    lngOrder As Long
    intSize As Integer
    blnEnemy As Boolean
End Type
Global glngRadarOrder As Long
Global glngNumRadar As Long
Global gudtRadar() As RADARTYPE

'Distress call stuff
Global Const DISTRESS_RANGE = 1000000
Global gblnDistress As Boolean     'Has there been a distress call?
Global gdblDistressX As Double
Global gdblDistressY As Double

'AI action constants
Global Const ACTION_NONE = 0
Global Const ACTION_NOTHRUST = 1
Global Const ACTION_THRUST = 2
Global Const ACTION_REVERSETHRUST = 4
Global Const ACTION_NOTURN = 8
Global Const ACTION_LEFT = 16
Global Const ACTION_RIGHT = 32

'Seek algorithm constants
Global Const MIN_VECTOR_SPEED_DIFF = 0.005

'Shield constants
Global Const SHIELD_DURATION = 100
Global Const SHIELD_0 = 0
Global Const SHIELD_1 = 1
Global Const SHIELD_2 = 2
Global Const SHIELD_3 = 3
Global Const SHIELD_4 = 4
Global Const SHIELD_5 = 5
Global Const SHIELD_9 = 6

'Explosions
Global Const EXPLOSION_ANIM_RATE = 50   'ms per frame of explosion animation
Global Const EXPLOSION0_WIDTH = 50
Global Const EXPLOSION0_HEIGHT = 50
Global Const EXPLOSION1_WIDTH = 80
Global Const EXPLOSION1_HEIGHT = 80
Global Const EXPLOSION_NUM = 1
Global Const EXPLOSION_SPRITE_NUM = 8
Global glngNumExplosions As Long
Type EXPLOSIONSPRITETYPE
    lngSprite(EXPLOSION_SPRITE_NUM) As Long
End Type
Global gudtExplosionSprite(EXPLOSION_NUM) As EXPLOSIONSPRITETYPE
Type EXPLOSIONTYPE
    dblX As Double
    dblY As Double
    sngSpeed As Single
    sngDirection As Single
    bytAnimFrame As Byte
    lngNextFrameTime As Long
    bytExplosionType As Byte
End Type
Global gudtExplosion() As EXPLOSIONTYPE
'Explosion sounds
Global glngExplosionBig As Long     'References to our explosion sounds
Global glngExplosionSmall As Long

'Lasers
Global Const LASER_DISPLAY_DIST = 500
Global glngNumLaserDisplay As Long
Type LASERDISPLAYTYPE
    lngOwner As Long
    lngTarget As Long
    bytType As Byte
End Type
Global gudtLaserDisplay() As LASERDISPLAYTYPE

'Death/exploding variables
Global gblnPlayerDead As Boolean                        'Is the player certainly dead?
Global gblnPlayerExploded As Boolean                    'Has the player finished exploding?
Global gblnPlayerDeadMessage As Boolean                 'Has the "You're Dead" message been displayed?
Global glngPlayerExplodingStart As Long                 'When did the player start exploding?
Global glngPlayerExplosionNum As Long                   'How many explosions have occurred?
Global Const PLAYER_EXPLODE_CONTINUE_DELAY = 5000       'How long until the player is given the option "Press any key to return to main menu"?
Global Const PLAYER_EXPLODE_DURATION = 1200             'How long will the player explode for when he dies?
Global Const EXPLODE_DURATION = 600                     'How long will a non-player object explode for?
Global Const PIXELS_WIDTH_PER_EXPLOSION = 2             'How many small explosions per pixel of hull width?
Global Const PIXELS_WIDTH_FOR_LARGE_EXPLOSION = 40      'How many pixels of width are required to warrant a large ship explosion?
Type OBJECTDEADTYPE
    lngObjectNum As Long
    lngExplodingStart As Long
    lngExplosionNum As Long
End Type
Global gudtExplodingObject() As OBJECTDEADTYPE

'Sound play vars
Global Const MIN_SOUND_DIST = 500

'Randomness!
Global Const GREAT_DIST = 500000      'Distance at which AI should not consider firing
Global Const MAX_FRAME_LENGTH = 100   'Maximum number of MS per frame

'Number of milliseconds elapsed since game began
Global glngGameTime As Long

'Bit rolling constant
Global Const BIT_ROLL_CONSTANT = 60

Function BitRoll(bytData As Byte, lngSeed As Long, Optional blnInvert As Boolean = False) As Byte

Dim intRollAmount As Integer
    
    'Determine the size of the roll
    If lngSeed <= 0 Then lngSeed = Abs(lngSeed)
    intRollAmount = lngSeed Mod BIT_ROLL_CONSTANT
    If blnInvert Then intRollAmount = -intRollAmount

    'Roll the byte data over
    If bytData + intRollAmount > 255 Then
        BitRoll = bytData + intRollAmount - 256
    ElseIf bytData + intRollAmount < 0 Then
        BitRoll = bytData + intRollAmount + 256
    Else
        BitRoll = bytData + intRollAmount
    End If

End Function

Public Sub LoadUniverse(Optional blnLoadData As Boolean = True)

Dim bytData() As Byte
Dim bytDataTemp() As Byte
Dim i As Long

    'log
    Log "Universe", "LoadUniverse", "Loading the universe '" & gstrUniverse & "'..."

    'Load the universe file
    Open gstrUniverse For Binary Access Read Write Lock Write As #1
        'Is the file empty?
        If LOF(1) = 0 Then Exit Sub
        'Grab all the data
        ReDim bytData(LOF(1))
        ReDim bytDataTemp(LOF(1))
        Get 1, , bytData
        Get 1, 1, bytDataTemp
        'Un-bitroll it
        For i = 0 To UBound(bytData)
            bytData(i) = BitRoll(bytData(i), i, True)
        Next i
        'Put it back
        Put 1, 1, bytData
        'Now extract it array by array
        Get #1, 1, gudtArmour
        glngArmourNum = UBound(gudtArmour)
        Get #1, , gudtCannon
        Get #1, , gstrCannonSound
        Get #1, , gstrCannonSprite
        glngCannonNum = UBound(gudtCannon)
        Get #1, , gudtEngine
        Get #1, , gstrEngineSound
        glngEngineNum = UBound(gudtEngine)
        Get #1, , gudtGenerator
        glngGeneratorNum = UBound(gudtGenerator)
        Get #1, , gudtHull
        glngHullNum = UBound(gudtHull)
        Get #1, , gudtJammer
        Get #1, , gudtLaser
        Get #1, , gstrLaserSound
        glngLaserNum = UBound(gudtLaser)
        Get #1, , gudtMissile
        Get #1, , gstrMissileSound
        glngMissileNum = UBound(gudtMissile)
        Get #1, , gudtScanner
        glngScannerNum = UBound(gudtScanner)
        Get #1, , gudtShield
        glngShieldNum = UBound(gudtShield)
        Get #1, , gudtRace
        Dim lngNumObjects As Long
        Get #1, , lngNumObjects
        ReDim gudtObject(lngNumObjects)
        Get #1, , gudtObject
        Get #1, , glngGameTime
        Get #1, , gudtPlayer
        gudtPlayer.lngRadarObject = -1  'Ensure that we don't get a crash, from having an object select which doesn't exist!!  Remember "0" suggests that gudtObject(0) is in radar range.
        gudtPlayer.udtAI.lngTarget = TARGET_NONE
        'Put the bitrolled data back
        Put #1, 1, bytDataTemp
    Close #1

    'Load the rest of the necessary data
    If blnLoadData Then LoadData
    
    'The universe is loaded
    Log "Universe", "LoadUniverse", "Universe loaded!"
    gblnUniverseLoaded = True

End Sub

Public Sub LoadData()

Dim i As Long
Dim j As Long

    'log
    Log "Universe", "LoadData", "Loading the universe data..."

    'Init radar
    glngNumRadar = 0
    Erase gudtRadar
    glngRadarOrder = 0

    'Load the sprites
    Log "Universe", "LoadData", "Loading universe sprites"
    'Player Sprites
    gudtPlayer.udtSprite.lngSpriteObject = DDraw.LoadSpriteObject(gudtPlayer.udtSprite.strResName, gudtPlayer.udtSprite.intWidth, gudtPlayer.udtSprite.intHeight, gudtPlayer.udtSprite.bytFrameAmt, gudtPlayer.udtSprite.bytAnimAmt, True)  'The player sprite
    'Other ship sprites
    For i = 0 To UBound(gudtObject)
        'Load hull specific object info
        gudtObject(i).udtSprite.strResName = gudtHull(gudtObject(i).udtSystems.bytHull).udtSprite.strResName
        gudtObject(i).udtSprite.intWidth = gudtHull(gudtObject(i).udtSystems.bytHull).udtSprite.intWidth
        gudtObject(i).udtSprite.intHeight = gudtHull(gudtObject(i).udtSystems.bytHull).udtSprite.intHeight
        gudtObject(i).udtSprite.bytFrameAmt = gudtHull(gudtObject(i).udtSystems.bytHull).udtSprite.bytFrameAmt
        gudtObject(i).udtSprite.bytAnimAmt = gudtHull(gudtObject(i).udtSystems.bytHull).udtSprite.bytAnimAmt
        gudtObject(i).udtSprite.lngAnimRate = gudtHull(gudtObject(i).udtSystems.bytHull).udtSprite.lngAnimRate
        gudtObject(i).udtPhysics.lngMass = gudtHull(gudtObject(i).udtSystems.bytHull).lngMass
        gudtObject(i).udtSystems.sngRotationRate = gudtHull(gudtObject(i).udtSystems.bytHull).sngRotationRate
        'If GetDist(gudtObject(i).udtPhysics.dblX, gudtObject(i).udtPhysics.dblY, gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY) <= LOAD_DISTANCE And gudtObject(i).udtSprite.blnLoaded = False Then
            gudtObject(i).udtSprite.lngSpriteObject = DDraw.LoadSpriteObject(gudtObject(i).udtSprite.strResName, gudtObject(i).udtSprite.intWidth, gudtObject(i).udtSprite.intHeight, gudtObject(i).udtSprite.bytFrameAmt, gudtObject(i).udtSprite.bytAnimAmt, True)
            gudtObject(i).udtSprite.blnLoaded = True
        'End If
    Next i
    'Cannon sprites
    For i = 0 To CANNON_NUM
        gudtCannon(i).lngSprite = DDraw.LoadSprite(gstrCannonSprite(i), BULLET_WIDTH, BULLET_HEIGHT, 0)
        gudtCannon(i).lngSound = DSound.LoadSound(gstrCannonSound(i))
    Next i
    'Missile sprites
    For i = 0 To MISSILE_NUM
        gudtMissile(i).udtSprite.lngSpriteObject = DDraw.LoadSpriteObject(gudtMissile(i).udtSprite.strResName, gudtMissile(i).udtSprite.intWidth, gudtMissile(i).udtSprite.intHeight, gudtMissile(i).udtSprite.bytFrameAmt, gudtMissile(i).udtSprite.bytAnimAmt, True, False)
        gudtMissile(i).lngSound = DSound.LoadSound(gstrMissileSound(i))
    Next i
    'Explosion sprites
    Log "Universe", "LoadData", "Loading the explosion sprites"
    For i = 0 To EXPLOSION_SPRITE_NUM
        gudtExplosionSprite(0).lngSprite(i) = DDraw.LoadSprite("0explode" & i, EXPLOSION0_WIDTH, EXPLOSION0_HEIGHT, 0)
    Next i
    For i = 0 To EXPLOSION_SPRITE_NUM
        gudtExplosionSprite(1).lngSprite(i) = DDraw.LoadSprite("1explode" & i, EXPLOSION1_WIDTH, EXPLOSION1_HEIGHT, 0)
    Next i
    
    'Load engine sounds
    For i = 0 To ENGINE_NUM
        gudtEngine(i).lngSound = DSound.LoadSound(gstrEngineSound(i))
    Next i
    'Load laser sounds
    For i = 0 To LASER_NUM
        gudtLaser(i).lngSound = DSound.LoadSound(gstrLaserSound(i))
    Next i
    'Load explosion sounds
    Log "Universe", "LoadData", "Loading explosion sounds"
    glngExplosionBig = DSound.LoadSound("ExplodeBig")
    glngExplosionSmall = DSound.LoadSound("ExplodeSmall")
    
    'Reset the timer
    glngTimer = gobjDX.TickCount()
    
End Sub

Public Sub SaveUniverse()

Dim bytData() As Byte
Dim i As Long

    'log
    Log "Universe", "SaveUniverse", "Saving the universe '" & gstrUniverse & "'..."

    'If the file exists, delete it
    KillFile gstrUniverse
    
    Open gstrUniverse For Binary Access Read Write Lock Write As #1
        'Put the data in the file
        Put #1, 1, gudtArmour
        Put #1, , gudtCannon
        Put #1, , gstrCannonSound
        Put #1, , gstrCannonSprite
        Put #1, , gudtEngine
        Put #1, , gstrEngineSound
        Put #1, , gudtGenerator
        Put #1, , gudtHull
        Put #1, , gudtJammer
        Put #1, , gudtLaser
        Put #1, , gstrLaserSound
        Put #1, , gudtMissile
        Put #1, , gstrMissileSound
        Put #1, , gudtScanner
        Put #1, , gudtShield
        Put #1, , gudtRace
        Dim lngNumObjects As Long
        lngNumObjects = UBound(gudtObject)
        Put #1, , lngNumObjects
        Put #1, , gudtObject
        Put #1, , glngGameTime
        Put #1, , gudtPlayer
        'Put the bitrolled data back
        ReDim bytData(LOF(1))
        Get #1, 1, bytData
        For i = 0 To UBound(bytData)
            bytData(i) = BitRoll(bytData(i), i)
        Next i
        Put #1, 1, bytData
    Close #1

End Sub

Public Sub Load()

    'THIS FUNCTION IS DEPRECATED!
    'USE ONLY TO REGENERATE A UNIVERSE FILE OR SOMETHING

Dim i As Long
Dim j As Long

    'log
    Log "Universe", "Load", "Loading the universe..."

    'Init radar
    glngNumRadar = 0
    Erase gudtRadar
    glngRadarOrder = 0

    'Load explosion sprites
    Log "Universe", "Load", "Loading the explosion sprites"
    For i = 0 To EXPLOSION_NUM
        'glngSpriteExplosion(i) = DDraw.LoadSprite("explode" & i, EXPLOSION_WIDTH, EXPLOSION_HEIGHT, 0)
    Next i
    'Load explosion sounds
    Log "Universe", "Load", "Loading explosion sounds"
    glngExplosionBig = DSound.LoadSound("ExplodeBig")
    glngExplosionSmall = DSound.LoadSound("ExplodeSmall")

    'Load the default universe!
    Log "Universe", "Load", "Populating universe"
    If gstrUniverse = DEFAULT_UNIVERSE Then
        'Load the engines
        For i = 0 To ENGINE_NUM
            gudtEngine(i).sngThrust = i * 0.075
            gudtEngine(i).lngMaxEnergy = i * 20
            gudtEngine(i).sngConsumption = i * 0.001
            gudtEngine(i).strName = "Engine " & i
            gudtEngine(i).lngSound = DSound.LoadSound("Engine4")
        Next i
        'Load the generators
        For i = 0 To GENERATOR_NUM
            gudtGenerator(i).sngOutPut = i * 0.005
            gudtGenerator(i).lngMaxEnergy = i * 20
            gudtGenerator(i).lngMaxBattery = i * 50
            gudtGenerator(i).sngConsumption = i * 0.001
            gudtGenerator(i).strName = "Generator " & i
        Next i
        'Load the shields
        For i = 0 To SHIELD_NUM
            gudtShield(i).lngMaxEnergy = i * 20
            gudtShield(i).sngAbsorbtion = i
            gudtShield(i).sngConsumption = i * 0.001
            gudtShield(i).strName = "Shield " & i
        Next i
        'Load the lasers
        For i = 0 To LASER_NUM
            gudtLaser(i).lngMaxEnergy = i * 20
            gudtLaser(i).lngRange = 230
            gudtLaser(i).sngConcussiveDamage = 0.015 * i
            gudtLaser(i).sngRadiationDamage = 0.04 * i
            gudtLaser(i).sngConsumption = i * 0.0005
            gudtLaser(i).sngFireConsumption = i * 0.035
            gudtLaser(i).strName = "Laser " & i
            gudtLaser(i).lngColour = vbRed
            gudtLaser(i).lngSound = DSound.LoadSound("Laser1")
        Next i
        'Load the cannons
        For i = 0 To CANNON_NUM
            gudtCannon(i).lngDuration = 500 + i * 100
            gudtCannon(i).lngMaxEnergy = i * 20
            gudtCannon(i).sngConcussiveDamage = 2.5 * i
            gudtCannon(i).sngRadiationDamage = 0.5 * i
            gudtCannon(i).sngConsumption = i * 0.002
            gudtCannon(i).sngSpeed = 0.3 + 0.05 * i
            gudtCannon(i).strName = "Cannon " & i
            gudtCannon(i).lngFireRate = 500 - i * 75
            gudtCannon(i).lngSprite = DDraw.LoadSprite("Cannon0", BULLET_WIDTH, BULLET_HEIGHT, 0)
            gudtCannon(i).lngDuration = 500
            gudtCannon(i).sngInstantaneousConsumption = i * 0.5
            gudtCannon(i).lngSound = DSound.LoadSound("Cannon1")
        Next i
        'Load the missiles
        For i = 0 To MISSILE_NUM
            gudtMissile(i).lngDuration = 5000
            gudtMissile(i).sngConcussiveDamage = 20 * i
            gudtMissile(i).sngMaxSpeed = 0.99 * NORMALIZE_SPEED
            gudtMissile(i).sngRadiationDamage = 5 * i
            gudtMissile(i).sngRotationRate = 0.03
            gudtMissile(i).lngFireRate = 2000
            gudtMissile(i).sngSeekDist = 10
            gudtMissile(i).sngTargetBias = 0
            gudtMissile(i).sngThrust = 0.003
            gudtMissile(i).strName = "Missile " & i
            gudtMissile(i).udtSprite.blnLoaded = False
            gudtMissile(i).udtSprite.bytAnimAmt = 0
            gudtMissile(i).udtSprite.bytFrameAmt = FRAME_NUM
            gudtMissile(i).udtSprite.intHeight = 5
            gudtMissile(i).udtSprite.intWidth = 5
            gudtMissile(i).udtSprite.strResName = "1Missile"
            gudtMissile(i).udtSprite.lngSpriteObject = DDraw.LoadSpriteObject(gudtMissile(i).udtSprite.strResName, gudtMissile(i).udtSprite.intWidth, gudtMissile(i).udtSprite.intHeight, gudtMissile(i).udtSprite.bytFrameAmt, gudtMissile(i).udtSprite.bytAnimAmt, True, False)
            gudtMissile(i).lngSound = DSound.LoadSound("Missile1")
        Next i
        'Load armour
        For i = 0 To ARMOUR_NUM
            gudtArmour(i).strName = "Armour " & i
            gudtArmour(i).lngMaxArmour = i * 50
        Next i
        gudtArmour(1).lngMaxArmour = 25
        gudtArmour(5).lngMaxArmour = 500
        'Load scanners
        For i = 0 To SCANNER_NUM
            gudtScanner(i).strName = "Scanner " & i
            gudtScanner(i).dblMaxRange = MIN_SCANNER_RANGE + 2500 * i
        Next i
        'Load the hulls
        For i = 0 To HULL_NUM
            gudtHull(i).blnARCD = True
            gudtHull(i).blnCommJammer = True
            gudtHull(i).blnFTLD = True
            gudtHull(i).blnMines = True
            gudtHull(i).lngArmour = i * 25
            gudtHull(i).lngCannon = CANNON_NUM
            gudtHull(i).lngEngine = ENGINE_NUM
            gudtHull(i).lngGenerator = GENERATOR_NUM
            gudtHull(i).lngLaser = LASER_NUM
            gudtHull(i).lngMass = i * 100
            gudtHull(i).lngMaxCargo = i * 10
            gudtHull(i).lngMaxCrew = i * 30
            gudtHull(i).lngMaxMines = i
            gudtHull(i).lngMaxMissile = i * 2
            gudtHull(i).lngMissile = MISSILE_NUM
            gudtHull(i).lngMaxFuel = i * 2000
            gudtHull(i).lngShield = SHIELD_NUM
            gudtHull(i).strName = "Hull " & i
            'gudtHull(i).strResourceName = "Hull " & i
            gudtHull(i).sngRotationRate = i * 0.002
            gudtHull(i).sngMaxSpeed = 0.5 * NORMALIZE_SPEED
        Next i
        gudtHull(1).sngMaxSpeed = 0.95 * NORMALIZE_SPEED
        gudtHull(5).sngMaxSpeed = 0.75 * NORMALIZE_SPEED
        gudtHull(6).sngMaxSpeed = 0.6 * NORMALIZE_SPEED
        gudtHull(6).lngMaxCrew = 500
        
        'Jammer
        gudtJammer.sngConsumption = 0.025
        gudtJammer.dblRange = 10# * NORMALIZE_DISTANCE_AU
        
        'Load races
        gudtRace(RACE_KALE).blnEncountered = True
        gudtRace(RACE_KALE).intRelations(RACE_PLAYER) = -200
        gudtRace(RACE_KALE).intRelations(RACE_GRAME) = -200
        gudtRace(RACE_GRAME).intRelations(RACE_PLAYER) = 200
        gudtRace(RACE_GRAME).intRelations(RACE_KALE) = -200
        gudtRace(RACE_ALTAIRIAN).blnEncountered = True
        gudtRace(RACE_ALTAIRIAN).intRelations(RACE_PLAYER) = 200
        gudtRace(RACE_VEGAN).blnEncountered = True
        gudtRace(RACE_VEGAN).intRelations(RACE_PLAYER) = 0
        gudtRace(RACE_GRAME).blnEncountered = True
        gudtRace(RACE_GRAME).intRelations(RACE_SICARIUS) = -200
        gudtRace(RACE_TERRAN).intRelations(RACE_PLAYER) = 0
        gudtRace(RACE_TERRAN).intRelations(RACE_KALE) = -200
        gudtRace(RACE_TERRAN).intRelations(RACE_SICARIUS) = -200
        gudtRace(RACE_SICARIUS).intRelations(RACE_PLAYER) = -200
        gudtRace(RACE_SICARIUS).intRelations(RACE_TERRAN) = -200
        gudtRace(RACE_TERRAN).strName = "Terran"
        gudtRace(RACE_KALE).strName = "Kale"
        gudtRace(RACE_PRAEMALI).strName = "Praemali"
        gudtRace(RACE_HANTAKAS).strName = "Hantakas"
        gudtRace(RACE_ALTAIRIAN).strName = "Altairian"
        gudtRace(RACE_GRAME).strName = "Grame"
        gudtRace(RACE_VEGAN).strName = "Vegan"
        gudtRace(RACE_ULWAR).strName = "Ulwar"
        gudtRace(RACE_TULONI).strName = "Tuloni"
        gudtRace(RACE_SICARIUS).strName = "Sicarius"
        gudtRace(RACE_INDEPENDENT).strName = "Independent"
        gudtRace(RACE_PLAYER).strName = "Player"
        gudtRace(RACE_PLANET).strName = "Planet"
        
        'Load the player!
        gudtPlayer.udtPhysics.dblX = 0
        gudtPlayer.udtPhysics.dblY = 0
        gudtPlayer.udtSprite.blnLoaded = True
        gudtPlayer.udtSprite.bytAnimAmt = 0
        gudtPlayer.udtSprite.bytFrameAmt = FRAME_NUM
        gudtPlayer.udtSprite.intHeight = 40
        gudtPlayer.udtSprite.intWidth = 40
        gudtPlayer.udtSprite.strResName = "2Grame"
        gudtPlayer.udtSprite.lngAnimRate = 25
        gudtPlayer.udtSystems.bytEngine = 2
        gudtPlayer.udtSystems.sngEngineEnergy = gudtEngine(3).lngMaxEnergy
        gudtPlayer.udtSystems.bytHull = 5
        gudtPlayer.udtSystems.sngRotationRate = gudtHull(gudtPlayer.udtSystems.bytHull).sngRotationRate
        gudtPlayer.udtSystems.lngCrew = gudtHull(gudtPlayer.udtSystems.bytHull).lngMaxCrew
        gudtPlayer.udtPhysics.lngMass = 1000
        gudtPlayer.udtControl.bytEngine = 125
        gudtPlayer.udtControl.bytGenerator = 125
        gudtPlayer.udtControl.bytShield = 125
        gudtPlayer.udtControl.bytWeapons = 125
        gudtPlayer.udtSystems.sngFuel = gudtHull(5).lngMaxFuel
        gudtPlayer.udtSystems.bytGenerator = 7
        gudtPlayer.udtSystems.sngGeneratorEnergy = gudtGenerator(6).lngMaxEnergy
        gudtPlayer.udtSystems.sngEnergy = 100
        gudtPlayer.udtSystems.bytCannon = 5
        gudtPlayer.udtSystems.bytLaser = 3
        gudtPlayer.udtSystems.bytMissile = 5
        gudtPlayer.udtSystems.intMissileNum = 10
        gudtPlayer.udtSystems.sngWeaponEnergy = gudtCannon(2).lngMaxEnergy + gudtLaser(2).lngMaxEnergy
        gudtPlayer.udtSystems.bytShield = 3
        gudtPlayer.udtSystems.sngShieldEnergy = gudtShield(3).lngMaxEnergy
        gudtPlayer.udtSystems.bytArmour = 4
        gudtPlayer.udtSystems.lngArmour = gudtArmour(4).lngMaxArmour
        gudtPlayer.udtCargo.lngNumCargo = 3
        ReDim gudtPlayer.udtCargo.lngAmount(2)
        gudtPlayer.udtCargo.lngAmount(0) = 5
        gudtPlayer.udtCargo.lngAmount(1) = 5
        gudtPlayer.udtCargo.lngAmount(2) = 25
        gudtPlayer.udtCargo.lngSalvage = 6
        gudtPlayer.udtAI.dblX = 59999990000#
        gudtPlayer.udtAI.dblY = 53939990000#
        gudtPlayer.udtAI.lngTarget = 1
        gudtPlayer.udtSystems.bytScanner = 3
        gudtPlayer.udtInfo.strName = "Skyball"
        gudtPlayer.udtInfo.bytRace = RACE_PLAYER
        gudtPlayer.udtSystems.blnFTLD = True
        gudtPlayer.udtPhysics.dblX = 0 '59999990000#
        gudtPlayer.udtPhysics.dblY = 0 '9999990000#
        gudtPlayer.udtSystems.blnJammer = True
        gudtPlayer.udtSystems.bytScanner = 5
        gudtPlayer.dblCurrentRange = MIN_SCANNER_RANGE
        gudtPlayer.lngRadarObject = -1
    
        'Load sprite
        gudtPlayer.udtSprite.lngSpriteObject = DDraw.LoadSpriteObject(gudtPlayer.udtSprite.strResName, gudtPlayer.udtSprite.intWidth, gudtPlayer.udtSprite.intHeight, gudtPlayer.udtSprite.bytFrameAmt, gudtPlayer.udtSprite.bytAnimAmt, True)
        
        'Load the other objects
        ReDim gudtObject(49)
        For i = 0 To 49
            'Load the target lock time
            gudtObject(i).udtAI.lngLengthTargetLock = 500
        Next i
        
        gudtObject(45).udtInfo.strName = "Skyball"
        gudtObject(45).blnExists = True
        gudtObject(45).udtPhysics.dblX = 100
        gudtObject(45).udtPhysics.dblY = 5000
        gudtObject(45).udtSprite.blnLoaded = False
        gudtObject(45).udtSprite.bytAnimAmt = 0
        gudtObject(45).udtSprite.bytFrameAmt = FRAME_NUM
        gudtObject(45).udtSprite.intHeight = 40
        gudtObject(45).udtSprite.intWidth = 40
        gudtObject(45).udtSprite.strResName = "2Grame"
        gudtObject(45).udtSprite.lngAnimRate = 25
        gudtObject(45).udtSystems.sngRotationRate = 0.01
        gudtObject(45).udtSystems.bytEngine = 3
        gudtObject(45).udtSystems.sngEngineEnergy = 40
        gudtObject(45).udtSystems.sngEnergy = 300
        gudtObject(45).udtPhysics.sngSpeed = 0.0001
        gudtObject(45).udtPhysics.blnThrusting = False
        gudtObject(45).udtPhysics.blnTurningLeft = False
        gudtObject(45).udtPhysics.sngFacing = 1
        gudtObject(45).udtPhysics.lngMass = 500
        gudtObject(45).udtAI.lngTarget = -2
        gudtObject(45).udtAI.bytAction = AI_ATTACK
        gudtObject(45).udtAI.sngTargetBias = 0
        gudtObject(45).udtAI.sngSeekDist = 100
        gudtObject(45).udtAI.sngAimTolerance = 0.2
        gudtObject(45).udtAI.sngMinDist = 100
        gudtObject(45).udtAI.sngCannonDist = 250
        gudtObject(45).udtSystems.bytHull = 6
        gudtObject(45).udtSystems.bytCannon = 5
        gudtObject(45).udtSystems.bytGenerator = 5
        gudtObject(45).udtSystems.lngCrew = gudtHull(gudtObject(45).udtSystems.bytHull).lngMaxCrew
        gudtObject(45).udtInfo.bytRace = RACE_GRAME
        gudtObject(45).udtSystems.bytArmour = 5
        gudtObject(45).udtSystems.bytLaser = 2
        gudtObject(45).udtSystems.lngArmour = gudtArmour(gudtObject(45).udtSystems.bytArmour).lngMaxArmour
        gudtObject(45).udtSystems.bytMissile = 5
        gudtObject(45).udtSystems.intMissileNum = 3
        
        gudtObject(46).udtInfo.strName = "Slope Oak"
        gudtObject(46).blnExists = True
        gudtObject(46).udtPhysics.dblX = -100
        gudtObject(46).udtPhysics.dblY = 4000
        gudtObject(46).udtSprite.blnLoaded = False
        gudtObject(46).udtSprite.bytAnimAmt = 0
        gudtObject(46).udtSprite.bytFrameAmt = FRAME_NUM
        gudtObject(46).udtSprite.intHeight = 100
        gudtObject(46).udtSprite.intWidth = 100
        gudtObject(46).udtSprite.strResName = "5Terran"
        gudtObject(46).udtSprite.lngAnimRate = 25
        gudtObject(46).udtSystems.sngRotationRate = 0.005
        gudtObject(46).udtSystems.bytEngine = 3
        gudtObject(46).udtSystems.sngEngineEnergy = 40
        gudtObject(46).udtSystems.sngEnergy = 300
        gudtObject(46).udtPhysics.sngSpeed = 0.0001
        gudtObject(46).udtPhysics.blnThrusting = False
        gudtObject(46).udtPhysics.blnTurningLeft = False
        gudtObject(46).udtPhysics.sngFacing = 1
        gudtObject(46).udtPhysics.lngMass = 1000
        gudtObject(46).udtAI.lngTarget = -2
        gudtObject(46).udtAI.bytAction = AI_ATTACK
        gudtObject(46).udtAI.sngTargetBias = 0
        gudtObject(46).udtAI.sngSeekDist = 0
        gudtObject(46).udtAI.sngAimTolerance = 0.2
        gudtObject(46).udtAI.sngMinDist = 150
        gudtObject(46).udtAI.sngCannonDist = 250
        gudtObject(46).udtSystems.bytHull = 10
        gudtObject(46).udtSystems.bytCannon = 5
        gudtObject(46).udtSystems.bytGenerator = 8
        gudtObject(46).udtSystems.lngCrew = gudtHull(gudtObject(46).udtSystems.bytHull).lngMaxCrew
        gudtObject(46).udtInfo.bytRace = RACE_TERRAN
        gudtObject(46).udtSystems.bytArmour = 7
        gudtObject(46).udtSystems.bytLaser = 2
        gudtObject(46).udtSystems.lngArmour = gudtArmour(gudtObject(46).udtSystems.bytArmour).lngMaxArmour
        gudtObject(46).udtSystems.bytMissile = 5
        gudtObject(46).udtSystems.intMissileNum = 3
        
        gudtObject(47).udtInfo.strName = "Flayer"
        gudtObject(47).blnExists = True
        gudtObject(47).udtPhysics.dblX = 100
        gudtObject(47).udtPhysics.dblY = -19000
        gudtObject(47).udtSprite.blnLoaded = False
        gudtObject(47).udtSprite.bytAnimAmt = 19
        gudtObject(47).udtSprite.bytFrameAmt = 0
        gudtObject(47).udtSprite.intHeight = 100
        gudtObject(47).udtSprite.intWidth = 100
        gudtObject(47).udtSprite.strResName = "5Sicar"
        gudtObject(47).udtSprite.lngAnimRate = 25
        gudtObject(47).udtSystems.sngRotationRate = 0
        gudtObject(47).udtSystems.bytEngine = 0
        gudtObject(47).udtSystems.sngEngineEnergy = 0
        gudtObject(47).udtSystems.sngEnergy = 3000
        gudtObject(47).udtPhysics.sngSpeed = 0
        gudtObject(47).udtPhysics.blnThrusting = False
        gudtObject(47).udtPhysics.blnTurningLeft = False
        gudtObject(47).udtPhysics.sngFacing = 0
        gudtObject(47).udtPhysics.lngMass = 1000
        gudtObject(47).udtAI.lngTarget = -2
        gudtObject(47).udtAI.bytAction = AI_ATTACK
        gudtObject(47).udtAI.sngTargetBias = 0
        gudtObject(47).udtAI.sngSeekDist = 0
        gudtObject(47).udtAI.sngAimTolerance = 0.2
        gudtObject(47).udtAI.sngMinDist = 150
        gudtObject(47).udtAI.sngCannonDist = 350
        gudtObject(47).udtSystems.bytHull = 10
        gudtObject(47).udtSystems.bytCannon = 5
        gudtObject(47).udtSystems.bytGenerator = 8
        gudtObject(47).udtSystems.lngCrew = gudtHull(gudtObject(47).udtSystems.bytHull).lngMaxCrew
        gudtObject(47).udtInfo.bytRace = RACE_SICARIUS
        gudtObject(47).udtSystems.bytArmour = 10
        gudtObject(47).udtSystems.bytLaser = 0
        gudtObject(47).udtSystems.lngArmour = gudtArmour(gudtObject(47).udtSystems.bytArmour).lngMaxArmour
        gudtObject(47).udtSystems.bytMissile = 4
        gudtObject(47).udtSystems.intMissileNum = 10
        
        gudtObject(48).udtInfo.strName = "Dalarak"
        gudtObject(48).blnExists = True
        gudtObject(48).udtPhysics.dblX = -1000
        gudtObject(48).udtPhysics.dblY = -11400 '-108540
        gudtObject(48).udtSprite.blnLoaded = False
        gudtObject(48).udtSprite.bytAnimAmt = 2
        gudtObject(48).udtSprite.bytFrameAmt = FRAME_NUM
        gudtObject(48).udtSprite.intHeight = 80
        gudtObject(48).udtSprite.intWidth = 80
        gudtObject(48).udtSprite.strResName = "4Kale"
        gudtObject(48).udtSprite.lngAnimRate = 100
        gudtObject(48).udtSystems.sngRotationRate = 0.01
        gudtObject(48).udtSystems.bytEngine = 2
        gudtObject(48).udtSystems.sngEngineEnergy = 40
        gudtObject(48).udtSystems.sngEnergy = 1000
        gudtObject(48).udtPhysics.sngSpeed = 0.0001
        gudtObject(48).udtPhysics.blnThrusting = False
        gudtObject(48).udtPhysics.blnTurningLeft = False
        gudtObject(48).udtPhysics.sngFacing = 1
        gudtObject(48).udtPhysics.lngMass = 1000
        gudtObject(48).udtAI.lngTarget = -1
        gudtObject(48).udtAI.bytAction = AI_ATTACK
        gudtObject(48).udtAI.sngTargetBias = 0
        gudtObject(48).udtAI.sngSeekDist = -100
        gudtObject(48).udtAI.sngAimTolerance = 0.2
        gudtObject(48).udtAI.sngMinDist = 10
        gudtObject(48).udtAI.sngCannonDist = 250
        gudtObject(48).udtSystems.bytHull = 4
        gudtObject(48).udtSystems.bytCannon = 3
        gudtObject(48).udtSystems.bytGenerator = 5
        gudtObject(48).udtSystems.lngCrew = gudtHull(gudtObject(48).udtSystems.bytHull).lngMaxCrew
        gudtObject(48).udtInfo.bytRace = RACE_KALE
        gudtObject(48).udtSystems.bytArmour = 5
        gudtObject(48).udtSystems.lngArmour = gudtArmour(gudtObject(48).udtSystems.bytArmour).lngMaxArmour
        gudtObject(48).udtSystems.bytLaser = 2
        gudtObject(48).udtSystems.bytMissile = 4
        gudtObject(48).udtSystems.intMissileNum = 3
        
        gudtObject(49).udtInfo.strName = "Slope Oak"
        gudtObject(49).blnExists = True
        gudtObject(49).udtPhysics.dblX = 800
        gudtObject(49).udtPhysics.dblY = 5000
        gudtObject(49).udtSprite.blnLoaded = False
        gudtObject(49).udtSprite.bytAnimAmt = 0
        gudtObject(49).udtSprite.bytFrameAmt = FRAME_NUM
        gudtObject(49).udtSprite.intHeight = 100
        gudtObject(49).udtSprite.intWidth = 100
        gudtObject(49).udtSprite.strResName = "5Terran"
        gudtObject(49).udtSprite.lngAnimRate = 25
        gudtObject(49).udtSystems.sngRotationRate = 0.005
        gudtObject(49).udtSystems.bytEngine = 3
        gudtObject(49).udtSystems.sngEngineEnergy = 40
        gudtObject(49).udtSystems.sngEnergy = 300
        gudtObject(49).udtPhysics.sngSpeed = 0.0001
        gudtObject(49).udtPhysics.blnThrusting = False
        gudtObject(49).udtPhysics.blnTurningLeft = False
        gudtObject(49).udtPhysics.sngFacing = 1
        gudtObject(49).udtPhysics.lngMass = 1000
        gudtObject(49).udtAI.lngTarget = -2
        gudtObject(49).udtAI.bytAction = AI_ATTACK
        gudtObject(49).udtAI.sngTargetBias = 0
        gudtObject(49).udtAI.sngSeekDist = -200
        gudtObject(49).udtAI.sngAimTolerance = 0.2
        gudtObject(49).udtAI.sngMinDist = 250
        gudtObject(49).udtAI.sngCannonDist = 250
        gudtObject(49).udtSystems.bytHull = 10
        gudtObject(49).udtSystems.bytCannon = 5
        gudtObject(49).udtSystems.bytGenerator = 8
        gudtObject(49).udtSystems.lngCrew = gudtHull(gudtObject(49).udtSystems.bytHull).lngMaxCrew
        gudtObject(49).udtInfo.bytRace = RACE_TERRAN
        gudtObject(49).udtSystems.bytArmour = 7
        gudtObject(49).udtSystems.bytLaser = 2
        gudtObject(49).udtSystems.lngArmour = gudtArmour(gudtObject(49).udtSystems.bytArmour).lngMaxArmour
        gudtObject(49).udtSystems.bytMissile = 5
        gudtObject(49).udtSystems.intMissileNum = 5
        
        gudtObject(0).udtInfo.strName = "Ankalan"
        gudtObject(0).blnExists = True
        gudtObject(0).udtPhysics.dblX = 100
        gudtObject(0).udtPhysics.dblY = -10000 '-108540
        gudtObject(0).udtSprite.blnLoaded = False
        gudtObject(0).udtSprite.bytAnimAmt = 2
        gudtObject(0).udtSprite.bytFrameAmt = FRAME_NUM
        gudtObject(0).udtSprite.intHeight = 80
        gudtObject(0).udtSprite.intWidth = 80
        gudtObject(0).udtSprite.strResName = "4Kale"
        gudtObject(0).udtSprite.lngAnimRate = 100
        gudtObject(0).udtSystems.sngRotationRate = 0.01
        gudtObject(0).udtSystems.bytEngine = 2
        gudtObject(0).udtSystems.sngEngineEnergy = 40
        gudtObject(0).udtSystems.sngEnergy = 1000
        gudtObject(0).udtPhysics.sngSpeed = 0.0001
        gudtObject(0).udtPhysics.blnThrusting = False
        gudtObject(0).udtPhysics.blnTurningLeft = False
        gudtObject(0).udtPhysics.sngFacing = 1
        gudtObject(0).udtPhysics.lngMass = 1000
        gudtObject(0).udtAI.lngTarget = -1
        gudtObject(0).udtAI.bytAction = AI_ATTACK
        gudtObject(0).udtAI.sngTargetBias = 0
        gudtObject(0).udtAI.sngSeekDist = 50
        gudtObject(0).udtAI.sngAimTolerance = 0.2
        gudtObject(0).udtAI.sngMinDist = 100
        gudtObject(0).udtAI.sngCannonDist = 150
        gudtObject(0).udtSystems.bytHull = 4
        gudtObject(0).udtSystems.bytCannon = 3
        gudtObject(0).udtSystems.bytGenerator = 5
        gudtObject(0).udtSystems.lngCrew = gudtHull(gudtObject(0).udtSystems.bytHull).lngMaxCrew
        gudtObject(0).udtInfo.bytRace = RACE_KALE
        gudtObject(0).udtSystems.bytArmour = 5
        gudtObject(0).udtSystems.lngArmour = gudtArmour(gudtObject(0).udtSystems.bytArmour).lngMaxArmour
        gudtObject(0).udtSystems.bytLaser = 2
        gudtObject(0).udtSystems.bytMissile = 4
        gudtObject(0).udtSystems.intMissileNum = 3
        
        gudtObject(1).blnExists = True
        gudtObject(1).udtPhysics.dblX = 400
        gudtObject(1).udtPhysics.dblY = 0
        gudtObject(1).udtSprite.blnLoaded = False
        gudtObject(1).udtSprite.bytAnimAmt = 4
        gudtObject(1).udtSprite.bytFrameAmt = 0
        gudtObject(1).udtSprite.intHeight = 300
        gudtObject(1).udtSprite.intWidth = 400
        gudtObject(1).udtSprite.strResName = "9Altair"
        gudtObject(1).udtSprite.lngAnimRate = 100
        gudtObject(1).udtSystems.sngRotationRate = 0
        gudtObject(1).udtSystems.bytEngine = 0
        gudtObject(1).udtSystems.sngEngineEnergy = 0
        gudtObject(1).udtPhysics.sngSpeed = 0
        gudtObject(1).udtInfo.strName = "Imperium"
        gudtObject(1).udtInfo.bytRace = RACE_ALTAIRIAN
        gudtObject(1).udtSystems.bytGenerator = 5
        gudtObject(1).udtSystems.sngEnergy = gudtGenerator(5).lngMaxBattery * 1000
        gudtObject(1).udtSystems.bytArmour = 5
        gudtObject(1).udtSystems.lngArmour = gudtArmour(gudtObject(1).udtSystems.bytArmour).lngMaxArmour * 3000
        
        gudtObject(2).blnExists = True
        gudtObject(2).udtPhysics.dblX = -2000
        gudtObject(2).udtPhysics.dblY = -250
        gudtObject(2).udtSprite.blnLoaded = False
        gudtObject(2).udtSprite.bytAnimAmt = 0
        gudtObject(2).udtSprite.bytFrameAmt = 0
        gudtObject(2).udtSprite.intHeight = 500
        gudtObject(2).udtSprite.intWidth = 500
        gudtObject(2).udtSprite.strResName = "Earth"
        gudtObject(2).udtSprite.lngAnimRate = 0
        gudtObject(2).udtSystems.sngRotationRate = 0
        gudtObject(2).udtInfo.bytRace = RACE_PLANET
        gudtObject(2).udtInfo.strName = "Terra"
        gudtObject(2).udtInfo.blnPlanet = True
        
        Randomize
        For i = 3 To 23
            gudtObject(i).blnExists = True
            gudtObject(i).udtSprite.blnLoaded = False
            gudtObject(i).udtSprite.bytFrameAmt = 39
            gudtObject(i).udtSprite.intHeight = 10
            gudtObject(i).udtSprite.intWidth = 10
            gudtObject(i).udtSprite.strResName = "0Kale"
            gudtObject(i).udtSystems.sngRotationRate = 0.01
            gudtObject(i).udtSystems.bytEngine = 1
            gudtObject(i).udtSystems.sngEngineEnergy = 20
            gudtObject(i).udtSystems.sngEnergy = gudtGenerator(1).lngMaxBattery
            gudtObject(i).udtPhysics.lngMass = 200
            gudtObject(i).udtAI.lngTarget = -1
            gudtObject(i).udtAI.bytAction = AI_ATTACK
            gudtObject(i).udtAI.sngTargetBias = 0.05 * Rnd() - 0.25 '* (i - 13)
            gudtObject(i).udtAI.sngSeekDist = -50
            gudtObject(i).udtAI.sngMinDist = 200
            
            gudtObject(i).udtAI.sngAimTolerance = 0.1
            gudtObject(i).udtAI.sngCannonDist = 150
            
            gudtObject(i).udtSystems.bytCannon = 1
            gudtObject(i).udtSystems.bytHull = 1
            gudtObject(i).udtSystems.bytGenerator = 1
            gudtObject(i).udtSystems.lngCrew = gudtHull(gudtObject(i).udtSystems.bytHull).lngMaxCrew
            gudtObject(i).udtSystems.bytArmour = 1
            gudtObject(i).udtSystems.lngArmour = gudtArmour(gudtObject(i).udtSystems.bytArmour).lngMaxArmour
            gudtObject(i).udtInfo.strName = "Subship " & i * 14
            gudtObject(i).udtInfo.bytRace = RACE_KALE
            
        Next i
        gudtObject(3).udtPhysics.dblX = 50000 / 2
        gudtObject(4).udtPhysics.dblX = 70000 / 2
        gudtObject(5).udtPhysics.dblX = 69000 / 2
        gudtObject(6).udtPhysics.dblY = 80000 / 2
        gudtObject(7).udtPhysics.dblY = 85000 / 2
        gudtObject(8).udtPhysics.dblY = 95000 / 2
        gudtObject(9).udtPhysics.dblY = 105000 / 2
        gudtObject(10).udtPhysics.dblY = 200000 / 2
        gudtObject(10).udtPhysics.dblX = 200000 / 2
        gudtObject(11).udtPhysics.dblY = -200000 / 2
        gudtObject(11).udtPhysics.dblX = 200000 / 2
        gudtObject(12).udtPhysics.dblY = 200000 / 2
        gudtObject(12).udtPhysics.dblX = -200000 / 2
        gudtObject(13).udtPhysics.dblY = -200000 / 2
        gudtObject(13).udtPhysics.dblX = -200000 / 2
        gudtObject(14).udtPhysics.dblY = 210000 / 2
        gudtObject(14).udtPhysics.dblX = 210000 / 2
        gudtObject(15).udtPhysics.dblY = -210000 / 2
        gudtObject(15).udtPhysics.dblX = 210000 / 2
        gudtObject(16).udtPhysics.dblY = 210000 / 2
        gudtObject(16).udtPhysics.dblX = -210000 / 2
        gudtObject(17).udtPhysics.dblY = -210000 / 2
        gudtObject(17).udtPhysics.dblX = -210000 / 2
        gudtObject(18).udtPhysics.dblY = 200000 / 2
        gudtObject(18).udtPhysics.dblX = 200000 / 2
        gudtObject(19).udtPhysics.dblY = -200000 / 2
        gudtObject(19).udtPhysics.dblX = 200000 / 2
        gudtObject(20).udtPhysics.dblY = 200000 / 2
        gudtObject(20).udtPhysics.dblX = -200000 / 2
        gudtObject(21).udtPhysics.dblY = -200000 / 2
        gudtObject(21).udtPhysics.dblX = -200000 / 2
        gudtObject(22).udtPhysics.dblY = 210000 / 2
        gudtObject(22).udtPhysics.dblX = 210000 / 2
        gudtObject(23).udtPhysics.dblY = -210000 / 2
        gudtObject(23).udtPhysics.dblX = 210000 / 2
        
        gudtObject(24).blnExists = True
        gudtObject(24).udtPhysics.dblX = 5000
        gudtObject(24).udtPhysics.dblY = -250
        gudtObject(24).udtSprite.blnLoaded = False
        gudtObject(24).udtSprite.bytAnimAmt = 0
        gudtObject(24).udtSprite.bytFrameAmt = 0
        gudtObject(24).udtSprite.intHeight = 500
        gudtObject(24).udtSprite.intWidth = 500
        gudtObject(24).udtSprite.strResName = "Star"
        gudtObject(24).udtSprite.lngAnimRate = 0
        gudtObject(24).udtSystems.sngRotationRate = 0
        gudtObject(24).udtInfo.bytRace = RACE_PLANET
        gudtObject(24).udtInfo.strName = "Altair"
        gudtObject(24).udtInfo.blnStar = True
        
        Randomize
        For i = 25 To 44
            gudtObject(i).blnExists = True
            gudtObject(i).udtSprite.blnLoaded = False
            gudtObject(i).udtSprite.bytFrameAmt = 39
            gudtObject(i).udtSprite.intHeight = 10
            gudtObject(i).udtSprite.intWidth = 10
            gudtObject(i).udtSprite.strResName = "0Kale"
            gudtObject(i).udtSystems.sngRotationRate = 0.01
            gudtObject(i).udtSystems.bytEngine = 1
            gudtObject(i).udtSystems.sngEngineEnergy = 20
            gudtObject(i).udtSystems.sngEnergy = gudtGenerator(2).lngMaxBattery
            gudtObject(i).udtPhysics.lngMass = 2000
            gudtObject(i).udtAI.lngTarget = -1
            gudtObject(i).udtAI.bytAction = AI_ATTACK
            gudtObject(i).udtAI.sngTargetBias = 0.02 * Rnd() - 0.1 '* (i - 25)
            gudtObject(i).udtAI.sngSeekDist = -50
            gudtObject(i).udtAI.sngMinDist = 100
            
            gudtObject(i).udtAI.sngAimTolerance = 0.1
            gudtObject(i).udtAI.sngCannonDist = 150
            
            gudtObject(i).udtSystems.bytCannon = 1
            gudtObject(i).udtSystems.bytHull = 1
            gudtObject(i).udtSystems.bytGenerator = 2
            gudtObject(i).udtSystems.lngCrew = gudtHull(gudtObject(i).udtSystems.bytHull).lngMaxCrew
            gudtObject(i).udtSystems.bytArmour = 1
            gudtObject(i).udtSystems.lngArmour = gudtArmour(gudtObject(i).udtSystems.bytArmour).lngMaxArmour
            gudtObject(i).udtInfo.strName = "Subship " & i * 14
            gudtObject(i).udtInfo.bytRace = RACE_KALE
            
            gudtObject(i).udtPhysics.dblX = 300000 - 5000 * Rnd() + 2500
            gudtObject(i).udtPhysics.dblY = 300000 - 5000 * Rnd() + 2500
        Next i
        
        'Load the sprites
        Log "Universe", "Load", "Loading universe sprites"
        For i = 0 To UBound(gudtObject)
            'If GetDist(gudtObject(i).udtPhysics.dblX, gudtObject(i).udtPhysics.dblY, gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY) <= LOAD_DISTANCE And gudtObject(i).udtSprite.blnLoaded = False Then
                gudtObject(i).udtSprite.lngSpriteObject = DDraw.LoadSpriteObject(gudtObject(i).udtSprite.strResName, gudtObject(i).udtSprite.intWidth, gudtObject(i).udtSprite.intHeight, gudtObject(i).udtSprite.bytFrameAmt, gudtObject(i).udtSprite.bytAnimAmt, True)
                gudtObject(i).udtSprite.blnLoaded = True
            'End If
        Next i
    End If
    
'    'Store all this muck in a universe file
'    Open "universe2.uni" For Binary Access Write As #1
'        Put #1, , gudtArmour
'        Put #1, , gudtCannon
'        Put #1, , gudtEngine
'        Put #1, , gudtGenerator
'        Put #1, , gudtHull
'        Put #1, , gudtJammer
'        Put #1, , gudtLaser
'        Put #1, , gudtMissile
'        Put #1, , gudtScanner
'        Put #1, , gudtShield
'
'        Put #1, , gudtRace
'        Dim glngNumObjects As Long
'        glngNumObjects = UBound(gudtObject)
'        Put #1, , glngNumObjects
'        Put #1, , gudtObject
'        Put #1, , gudtPlayer
'    Close #1
'
'    'Reset the timer
    glngTimer = gobjDX.TickCount()
    
    'The universe is loaded
    Log "Universe", "Load", "Universe loaded!"
    gblnUniverseLoaded = True
    
End Sub

Public Sub Physics()

Dim i As Long
Dim j As Long
Dim k As Long

    'Ensure we haven't passed max MS per frame
    If glngElapsed >= MAX_FRAME_LENGTH Then glngElapsed = MAX_FRAME_LENGTH

    'Explosions
    UpdateExplosions

    'Weapons
    BulletMove
    MissileMove
    
    'Move the player
    PlayerAI
    PlayerPhysics
    
    'Move the objects
    For i = 0 To UBound(gudtObject)
        'Ensure object exists
        If gudtObject(i).blnExists Then
            'Get the distance
            gudtObject(i).udtInfo.dblDistance = GetDist(gudtObject(i).udtPhysics.dblX, gudtObject(i).udtPhysics.dblY, gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY)
            'Ensure object within range are loaded
            If gudtObject(i).udtInfo.dblDistance <= LOAD_DISTANCE And gudtObject(i).udtSprite.blnLoaded = False Then
                'Queue (or load, if this is the first time)
                gudtObject(i).udtSprite.lngSpriteObject = DDraw.LoadSpriteObject(gudtObject(i).udtSprite.strResName, gudtObject(i).udtSprite.intWidth, gudtObject(i).udtSprite.intHeight, gudtObject(i).udtSprite.bytFrameAmt, gudtObject(i).udtSprite.bytAnimAmt, False)
                'Loaded!
                gudtObject(i).udtSprite.blnLoaded = True
            End If
'            'Ensure objects outside range are unloaded
'            If gudtObject(i).udtInfo.dblDistance >= UNLOAD_DISTANCE And gudtObject(i).udtSprite.blnLoaded = True Then
'                'Delete
'                DDraw.DeleteSpriteObject gudtObject(i).udtSprite.lngSpriteObject
'                'Unloaded!
'                gudtObject(i).udtSprite.blnLoaded = False
'            End If
            'Animate
            If gudtObject(i).udtSprite.bytAnimAmt > 0 Then
                gudtObject(i).udtSprite.lngAnimLast = gudtObject(i).udtSprite.lngAnimLast + glngElapsed
                If gudtObject(i).udtSprite.lngAnimLast > gudtObject(i).udtSprite.lngAnimRate Then
                    'Do it!
                    gudtObject(i).udtSprite.lngAnimLast = 0
                    gudtObject(i).udtSprite.bytAnimNum = gudtObject(i).udtSprite.bytAnimNum + 1
                    If gudtObject(i).udtSprite.bytAnimNum > gudtObject(i).udtSprite.bytAnimAmt Then gudtObject(i).udtSprite.bytAnimNum = 0
                End If
            End If
            'Energy
            If gudtObject(i).udtSystems.bytGenerator > 0 And gudtObject(i).udtSystems.sngEnergy < gudtGenerator(gudtObject(i).udtSystems.bytGenerator).lngMaxBattery Then
                gudtObject(i).udtSystems.sngEnergy = gudtObject(i).udtSystems.sngEnergy + gudtGenerator(gudtObject(i).udtSystems.bytGenerator).sngOutPut * glngElapsed
            End If
            'Turning
            If gudtObject(i).udtPhysics.blnTurningRight Then gudtObject(i).udtPhysics.sngFacing = FixAngle(gudtObject(i).udtPhysics.sngFacing + gudtObject(i).udtSystems.sngRotationRate * glngElapsed)
            If gudtObject(i).udtPhysics.blnTurningLeft Then gudtObject(i).udtPhysics.sngFacing = FixAngle(gudtObject(i).udtPhysics.sngFacing - gudtObject(i).udtSystems.sngRotationRate * glngElapsed)
            gudtObject(i).udtSprite.bytFrameNum = Fix((gudtObject(i).udtSprite.bytFrameAmt + 1) * (gudtObject(i).udtPhysics.sngFacing / (2 * Pi)))
            If gudtObject(i).udtSprite.bytFrameNum > FRAME_NUM Then gudtObject(i).udtSprite.bytFrameNum = FRAME_NUM
            'Calc mass
            gudtObject(i).udtPhysics.lngMass = gudtHull(gudtObject(i).udtSystems.bytHull).lngMass
            'If we're NOT FTL..
            If Not (gudtObject(i).udtSystems.blnARCDActive = True Or gudtObject(i).udtSystems.blnFTLDActive = True) Then
                'Thrusting
                If gudtObject(i).udtPhysics.blnThrusting Then AddVectors gudtObject(i).udtPhysics.sngSpeed, gudtObject(i).udtPhysics.sngHeading, gudtEngine(gudtObject(i).udtSystems.bytEngine).sngThrust * gudtObject(i).udtSystems.lngCrew / gudtHull(gudtObject(i).udtSystems.bytHull).lngMaxCrew * gudtObject(i).udtSystems.sngEngineEnergy / gudtEngine(gudtObject(i).udtSystems.bytEngine).lngMaxEnergy * glngElapsed / gudtObject(i).udtPhysics.lngMass, gudtObject(i).udtPhysics.sngFacing, gudtObject(i).udtPhysics.sngSpeed, gudtObject(i).udtPhysics.sngHeading
                If gudtObject(i).udtPhysics.blnReverseThrusting Then AddVectors gudtObject(i).udtPhysics.sngSpeed, gudtObject(i).udtPhysics.sngHeading, gudtEngine(gudtObject(i).udtSystems.bytEngine).sngThrust * gudtObject(i).udtSystems.lngCrew / gudtHull(gudtObject(i).udtSystems.bytHull).lngMaxCrew * gudtObject(i).udtSystems.sngEngineEnergy / gudtEngine(gudtObject(i).udtSystems.bytEngine).lngMaxEnergy * glngElapsed / gudtObject(i).udtPhysics.lngMass / 2, gudtObject(i).udtPhysics.sngFacing + Pi, gudtObject(i).udtPhysics.sngSpeed, gudtObject(i).udtPhysics.sngHeading
                'Cap speed if not using light drive..
                If gudtObject(i).udtPhysics.sngSpeed > gudtHull(gudtObject(i).udtSystems.bytHull).sngMaxSpeed Then gudtObject(i).udtPhysics.sngSpeed = gudtHull(gudtObject(i).udtSystems.bytHull).sngMaxSpeed
            Else
                'If we ARE FTL set speed
                If gudtObject(i).udtSystems.blnARCDActive = True Then gudtObject(i).udtPhysics.sngSpeed = ARCD_SPEED
                If gudtObject(i).udtSystems.blnFTLDActive = True Then gudtObject(i).udtPhysics.sngSpeed = FTLD_SPEED
                'Set direction
                gudtObject(i).udtPhysics.sngHeading = gudtObject(i).udtPhysics.sngFacing
            End If
            'Motion
            Motion gudtObject(i).udtPhysics.dblX, gudtObject(i).udtPhysics.dblY, gudtObject(i).udtPhysics.sngSpeed, gudtObject(i).udtPhysics.sngHeading
            'AI
            If gudtObject(i).udtAI.bytAction <> AI_NONE Then AI i
        End If
    Next i
    
    'Collision detection
    CheckCannonCollisions
    CheckMissileCollisions
        
End Sub

Private Sub MissileMove()

Dim bytAction As Byte
Dim sngDesiredFacing As Single
Dim sngTemp As Single
Dim i As Long

    'Check for missiles
    If glngNumLiveMissiles <= 0 Then Exit Sub
    
    'Recalc the missiles
    Do While i <= glngNumLiveMissiles - 1
        'Time to decay?
        If gudtLiveMissile(i).lngCreated + gudtMissile(gudtLiveMissile(i).bytMissile).lngDuration <= glngGameTime Then
            DeleteMissile i
            i = i - 1
        Else
            'Otherwise, move the missile
            If gudtLiveMissile(i).lngTarget = TARGET_PLAYER Then
                bytAction = SeekTargetNoRev(gudtLiveMissile(i).sngSpeed, gudtMissile(gudtLiveMissile(i).bytMissile).sngThrust, gudtLiveMissile(i).sngDirection, gudtLiveMissile(i).sngFacing, gudtLiveMissile(i).dblX, gudtLiveMissile(i).dblY, gudtPlayer.udtPhysics.sngSpeed, gudtPlayer.udtPhysics.sngHeading, gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, sngDesiredFacing, 0, gudtMissile(gudtLiveMissile(i).bytMissile).sngSeekDist, gudtMissile(gudtLiveMissile(i).bytMissile).sngTargetBias)
                sngTemp = FindAngle(gudtLiveMissile(i).dblX, gudtLiveMissile(i).dblY, gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY)
            ElseIf gudtLiveMissile(i).lngTarget <> TARGET_NONE Then
                bytAction = SeekTargetNoRev(gudtLiveMissile(i).sngSpeed, gudtMissile(gudtLiveMissile(i).bytMissile).sngThrust, gudtLiveMissile(i).sngDirection, gudtLiveMissile(i).sngFacing, gudtLiveMissile(i).dblX, gudtLiveMissile(i).dblY, gudtObject(gudtLiveMissile(i).lngTarget).udtPhysics.sngSpeed, gudtObject(gudtLiveMissile(i).lngTarget).udtPhysics.sngHeading, gudtObject(gudtLiveMissile(i).lngTarget).udtPhysics.dblX, gudtObject(gudtLiveMissile(i).lngTarget).udtPhysics.dblY, sngDesiredFacing, 0, gudtMissile(gudtLiveMissile(i).bytMissile).sngSeekDist, gudtMissile(gudtLiveMissile(i).bytMissile).sngTargetBias)
                sngTemp = FindAngle(gudtLiveMissile(i).dblX, gudtLiveMissile(i).dblY, gudtObject(gudtLiveMissile(i).lngTarget).udtPhysics.dblX, gudtObject(gudtLiveMissile(i).lngTarget).udtPhysics.dblY)
            End If
            'Take action
            'Fake the turning
            gudtLiveMissile(i).sngFacing = FixAngle(sngDesiredFacing)
            gudtLiveMissile(i).udtSprite.bytFrameNum = Fix((gudtMissile(gudtLiveMissile(i).bytMissile).udtSprite.bytFrameAmt + 1) * (FixAngle(sngTemp) / (2 * Pi)))
            If gudtLiveMissile(i).udtSprite.bytFrameNum > FRAME_NUM Then gudtLiveMissile(i).udtSprite.bytFrameNum = FRAME_NUM
            'Thrusting
            If (bytAction And ACTION_THRUST) = ACTION_THRUST Then AddVectors gudtLiveMissile(i).sngSpeed, gudtLiveMissile(i).sngDirection, gudtMissile(gudtLiveMissile(i).bytMissile).sngThrust * glngElapsed, gudtLiveMissile(i).sngFacing, gudtLiveMissile(i).sngSpeed, gudtLiveMissile(i).sngDirection
            If gudtLiveMissile(i).sngSpeed > gudtMissile(gudtLiveMissile(i).bytMissile).sngMaxSpeed Then gudtLiveMissile(i).sngSpeed = gudtMissile(gudtLiveMissile(i).bytMissile).sngMaxSpeed
            'Time to add to the trail?
            If gudtLiveMissile(i).lngSmokeTime + SMOKE_TRAIL_DELAY <= glngGameTime Then
                'Reset the clock
                gudtLiveMissile(i).lngSmokeTime = glngGameTime
                'Store the previous values
                gudtLiveMissile(i).dblXPrev(4) = gudtLiveMissile(i).dblXPrev(3)
                gudtLiveMissile(i).dblXPrev(3) = gudtLiveMissile(i).dblXPrev(2)
                gudtLiveMissile(i).dblXPrev(2) = gudtLiveMissile(i).dblXPrev(1)
                gudtLiveMissile(i).dblXPrev(1) = gudtLiveMissile(i).dblXPrev(0)
                gudtLiveMissile(i).dblXPrev(0) = gudtLiveMissile(i).dblX
                gudtLiveMissile(i).dblYPrev(4) = gudtLiveMissile(i).dblYPrev(3)
                gudtLiveMissile(i).dblYPrev(3) = gudtLiveMissile(i).dblYPrev(2)
                gudtLiveMissile(i).dblYPrev(2) = gudtLiveMissile(i).dblYPrev(1)
                gudtLiveMissile(i).dblYPrev(1) = gudtLiveMissile(i).dblYPrev(0)
                gudtLiveMissile(i).dblYPrev(0) = gudtLiveMissile(i).dblY
            End If
            'Moving
            Motion gudtLiveMissile(i).dblX, gudtLiveMissile(i).dblY, gudtLiveMissile(i).sngSpeed, gudtLiveMissile(i).sngDirection
        End If
        'Increment
        i = i + 1
    Loop

End Sub

Private Sub BulletMove()

Dim i As Long

    'Check for bullets
    If glngNumBullets <= 0 Then Exit Sub
    
    'Recalc the bullets
    Do While i <= glngNumBullets - 1
        'Time to decay?
        If gudtBullet(i).lngCreated + gudtCannon(gudtBullet(i).bytCannon).lngDuration <= glngGameTime Then
            DeleteBullet i
            i = i - 1
        Else
            'Otherwise move the bullet
            Motion gudtBullet(i).dblX, gudtBullet(i).dblY, gudtBullet(i).sngSpeed, gudtBullet(i).sngDirection
        End If
        'Increment..
        i = i + 1
    Loop

End Sub

Public Sub CreateBullet(bytCannon As Byte, dblX As Double, dblY As Double, lngOwner As Long, sngSpeed As Single, sngDirection As Single, sngFacing As Single)

Dim lngPan As Long
Dim lngVolume As Long

    'Create a new spot in the array
    ReDim Preserve gudtBullet(glngNumBullets)
    glngNumBullets = glngNumBullets + 1
    
    'Fill it with data
    With gudtBullet(glngNumBullets - 1)
        .bytCannon = bytCannon
        .dblX = dblX
        .dblY = dblY
        .lngCreated = glngGameTime
        .lngOwner = lngOwner
        AddVectors sngSpeed, sngDirection, gudtCannon(bytCannon).sngSpeed, sngFacing, .sngSpeed, .sngDirection
    End With
    
    'Get pan + vol
    GetPanAndVol dblX, dblY, lngPan, lngVolume
    
    'Play the appropriate sound
    DSound.PlaySound gudtCannon(bytCannon).lngSound, False, True, False, lngPan, lngVolume

End Sub

Public Function CreateMissile(bytMissile As Byte, dblX As Double, dblY As Double, lngOwner As Long, lngTarget As Long, sngSpeed As Single, sngDirection As Single, sngFacing As Single) As Boolean

Dim lngPan As Long
Dim lngVolume As Long

    'Not yet successful
    CreateMissile = False

    'Is this an appropriate missle type?
    If bytMissile = 0 Then Exit Function

    'Is the target a planet or star?
    If lngTarget <> -1 Then If gudtObject(lngTarget).udtInfo.bytRace = RACE_PLANET Then Exit Function

    'Create a new spot in the array
    ReDim Preserve gudtLiveMissile(glngNumLiveMissiles)
    glngNumLiveMissiles = glngNumLiveMissiles + 1
    
    'Fill it with data
    With gudtLiveMissile(glngNumLiveMissiles - 1)
        .bytMissile = bytMissile
        .dblX = dblX
        .dblY = dblY
        .lngCreated = glngGameTime
        .lngSmokeTime = glngGameTime
        .lngOwner = lngOwner
        .lngTarget = lngTarget
        .sngSpeed = sngSpeed
        .sngDirection = sngDirection
        .sngFacing = sngFacing
        .udtSprite.bytAnimAmt = gudtMissile(bytMissile).udtSprite.bytAnimAmt
        .udtSprite.bytFrameAmt = gudtMissile(bytMissile).udtSprite.bytFrameAmt
        .udtSprite.lngAnimRate = gudtMissile(bytMissile).udtSprite.lngAnimRate
        .udtSprite.lngSpriteObject = gudtMissile(bytMissile).udtSprite.lngSpriteObject
    End With
    
    'Get pan + vol
    GetPanAndVol dblX, dblY, lngPan, lngVolume
    
    'Play the appropriate sound
    DSound.PlaySound gudtMissile(bytMissile).lngSound, False, True, False, lngPan, lngVolume
    
    'Success!
    CreateMissile = True

End Function

Sub GetPanAndVol(dblX As Double, dblY As Double, ByRef lngPan As Long, ByRef lngVolume As Long)

Dim dblDist As Double
Dim lngDeltaX As Long
    
    'Calc dist
    dblDist = GetDist(gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, dblX, dblY)
    
    'If dist is too far, exit
    If dblDist > MIN_SOUND_DIST Then
        lngPan = 0
        lngVolume = VOL_MIN
        Exit Sub
    End If
    
    'Calc pan
    lngDeltaX = dblX - gudtPlayer.udtPhysics.dblX
    If lngDeltaX < 0 And lngDeltaX > -PAN_MAX_DIST Then
        lngPan = NormalizeLogScale(PAN_MAX_DIST - Abs(lngDeltaX), PAN_MAX_DIST, PAN_ATTENUATION, PAN_LEFT, PAN_CENTER)
    ElseIf lngDeltaX > 0 And lngDeltaX < PAN_MAX_DIST Then
        lngPan = NormalizeLogScale(PAN_MAX_DIST - lngDeltaX, PAN_MAX_DIST, PAN_ATTENUATION, PAN_RIGHT, PAN_CENTER)
    ElseIf lngDeltaX < -PAN_MAX_DIST Then
        lngPan = PAN_LEFT
    ElseIf lngDeltaX > PAN_MAX_DIST Then
        lngPan = PAN_RIGHT
    Else
        lngPan = 0
    End If
    
    'Calc vol (no sound at minDist, full at 0)
    lngVolume = NormalizeLogScale(MIN_SOUND_DIST - CLng(dblDist), MIN_SOUND_DIST, VOL_ATTENUATION, VOL_MIN, VOL_MAX)

End Sub

Private Sub DeleteMissile(lngMissile As Long)

Dim i As Long

    'Is there such a bullet?
    If lngMissile >= glngNumLiveMissiles Then Exit Sub
    
    'If this is the last missile, so be it
    If glngNumLiveMissiles = 1 Then
        Erase gudtLiveMissile
        glngNumLiveMissiles = 0
        Exit Sub
    End If
    
    'Otherwise, remove it and decrement!
    For i = lngMissile To glngNumLiveMissiles - 2
        gudtLiveMissile(i).bytMissile = gudtLiveMissile(i + 1).bytMissile
        gudtLiveMissile(i).dblX = gudtLiveMissile(i + 1).dblX
        gudtLiveMissile(i).dblY = gudtLiveMissile(i + 1).dblY
        gudtLiveMissile(i).lngCreated = gudtLiveMissile(i + 1).lngCreated
        gudtLiveMissile(i).lngOwner = gudtLiveMissile(i + 1).lngOwner
        gudtLiveMissile(i).lngTarget = gudtLiveMissile(i + 1).lngTarget
        gudtLiveMissile(i).sngDirection = gudtLiveMissile(i + 1).sngDirection
        gudtLiveMissile(i).sngFacing = gudtLiveMissile(i + 1).sngFacing
        gudtLiveMissile(i).sngSpeed = gudtLiveMissile(i + 1).sngSpeed
        gudtLiveMissile(i).udtSprite.bytAnimNum = gudtLiveMissile(i + 1).udtSprite.bytAnimNum
        gudtLiveMissile(i).udtSprite.bytFrameNum = gudtLiveMissile(i + 1).udtSprite.bytFrameNum
        gudtLiveMissile(i).udtSprite.lngAnimLast = gudtLiveMissile(i + 1).udtSprite.lngAnimLast
        gudtLiveMissile(i).udtSprite.lngAnimRate = gudtLiveMissile(i + 1).udtSprite.lngAnimRate
        gudtLiveMissile(i).udtSprite.lngSpriteObject = gudtLiveMissile(i + 1).udtSprite.lngSpriteObject
    Next i
    ReDim Preserve gudtLiveMissile(glngNumLiveMissiles - 2)
    glngNumLiveMissiles = glngNumLiveMissiles - 1

End Sub

Private Sub DeleteBullet(lngBullet As Long)

Dim i As Long

    'Is there such a bullet?
    If lngBullet >= glngNumBullets Then Exit Sub

    'If this is the last bullet, so be it
    If glngNumBullets = 1 Then
        Erase gudtBullet
        glngNumBullets = 0
        Exit Sub
    End If
    
    'Otherwise, remove and decrement!
    For i = lngBullet To glngNumBullets - 2
        gudtBullet(i).bytCannon = gudtBullet(i + 1).bytCannon
        gudtBullet(i).dblX = gudtBullet(i + 1).dblX
        gudtBullet(i).dblY = gudtBullet(i + 1).dblY
        gudtBullet(i).lngCreated = gudtBullet(i + 1).lngCreated
        gudtBullet(i).lngOwner = gudtBullet(i + 1).lngOwner
        gudtBullet(i).sngDirection = gudtBullet(i + 1).sngDirection
        gudtBullet(i).sngSpeed = gudtBullet(i + 1).sngSpeed
    Next i
    ReDim Preserve gudtBullet(glngNumBullets - 2)
    glngNumBullets = glngNumBullets - 1

End Sub

Public Sub PlayerFireCannon()

Dim lngTemp As Long
Dim sngHeading As Single
Dim dblX As Double
Dim dblY As Double

    'Do we HAVE cannons?
    If gudtPlayer.udtSystems.bytCannon <= 0 Then Exit Sub

    'Check crew..
    If gudtPlayer.udtSystems.lngCrew <= 0 Then Exit Sub
    
    'Check weapons energy
    If gudtPlayer.udtSystems.sngWeaponEnergy <= 0 Then Exit Sub

    'If it's a turret, and there's no target, exit sub
    If (gudtPlayer.lngRadarObject < 0) And (gudtCannon(gudtPlayer.udtSystems.bytCannon).lngCannonType And CANNON_TURRET = CANNON_TURRET) Then Exit Sub

    'Calculate duration of delay
    lngTemp = gudtCannon(gudtPlayer.udtSystems.bytCannon).lngFireRate / (gudtPlayer.udtSystems.lngCrew / gudtHull(gudtPlayer.udtSystems.bytHull).lngMaxCrew) / (gudtPlayer.udtSystems.sngWeaponEnergy / (gudtCannon(gudtPlayer.udtSystems.bytCannon).lngMaxEnergy + gudtLaser(gudtPlayer.udtSystems.bytLaser).lngMaxEnergy))
    
    'Is it time?
    If gudtPlayer.udtSystems.lngCannonLastFire + lngTemp <= glngGameTime Then
        'Do we have energy?
        If gudtPlayer.udtSystems.sngEnergy + gudtPlayer.udtSystems.sngWeaponEnergy >= gudtCannon(gudtPlayer.udtSystems.bytCannon).sngInstantaneousConsumption Then
            'Remove energy
            If gudtPlayer.udtSystems.sngEnergy >= gudtCannon(gudtPlayer.udtSystems.bytCannon).sngInstantaneousConsumption Then
                gudtPlayer.udtSystems.sngEnergy = gudtPlayer.udtSystems.sngEnergy - gudtCannon(gudtPlayer.udtSystems.bytCannon).sngInstantaneousConsumption
            Else
                gudtPlayer.udtSystems.sngWeaponEnergy = gudtPlayer.udtSystems.sngEnergy + gudtPlayer.udtSystems.sngWeaponEnergy
                gudtPlayer.udtSystems.sngWeaponEnergy = gudtPlayer.udtSystems.sngWeaponEnergy - gudtCannon(gudtPlayer.udtSystems.bytCannon).sngInstantaneousConsumption
                gudtPlayer.udtSystems.sngEnergy = 0
            End If
            'Store new lngCannonLastFire
            gudtPlayer.udtSystems.lngCannonLastFire = glngGameTime
            'Fire
            Select Case gudtCannon(gudtPlayer.udtSystems.bytCannon).lngCannonType
                Case (CANNON_FIXED Or CANNON_DOUBLE)
                    PointOnLine gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, gudtPlayer.udtPhysics.sngFacing + Pi / 2, CANNON_SPREAD, dblX, dblY
                    CreateBullet gudtPlayer.udtSystems.bytCannon, dblX, dblY, -1, gudtPlayer.udtPhysics.sngSpeed, gudtPlayer.udtPhysics.sngHeading, gudtPlayer.udtPhysics.sngFacing
                    PointOnLine gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, gudtPlayer.udtPhysics.sngFacing - Pi / 2, CANNON_SPREAD, dblX, dblY
                    CreateBullet gudtPlayer.udtSystems.bytCannon, dblX, dblY, -1, gudtPlayer.udtPhysics.sngSpeed, gudtPlayer.udtPhysics.sngHeading, gudtPlayer.udtPhysics.sngFacing
                Case (CANNON_FIXED Or CANNON_TRIPLE)
                    CreateBullet gudtPlayer.udtSystems.bytCannon, gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, -1, gudtPlayer.udtPhysics.sngSpeed, gudtPlayer.udtPhysics.sngHeading, gudtPlayer.udtPhysics.sngFacing
                    CreateBullet gudtPlayer.udtSystems.bytCannon, gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, -1, gudtPlayer.udtPhysics.sngSpeed, gudtPlayer.udtPhysics.sngHeading, gudtPlayer.udtPhysics.sngFacing + Pi / 9
                    CreateBullet gudtPlayer.udtSystems.bytCannon, gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, -1, gudtPlayer.udtPhysics.sngSpeed, gudtPlayer.udtPhysics.sngHeading, gudtPlayer.udtPhysics.sngFacing - Pi / 9
                Case (CANNON_FIXED Or CANNON_QUAD)
                    CreateBullet gudtPlayer.udtSystems.bytCannon, gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, -1, gudtPlayer.udtPhysics.sngSpeed, gudtPlayer.udtPhysics.sngHeading, gudtPlayer.udtPhysics.sngFacing
                    CreateBullet gudtPlayer.udtSystems.bytCannon, gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, -1, gudtPlayer.udtPhysics.sngSpeed, gudtPlayer.udtPhysics.sngHeading, gudtPlayer.udtPhysics.sngFacing + Pi / 2
                    CreateBullet gudtPlayer.udtSystems.bytCannon, gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, -1, gudtPlayer.udtPhysics.sngSpeed, gudtPlayer.udtPhysics.sngHeading, gudtPlayer.udtPhysics.sngFacing - Pi / 2
                    CreateBullet gudtPlayer.udtSystems.bytCannon, gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, -1, gudtPlayer.udtPhysics.sngSpeed, gudtPlayer.udtPhysics.sngHeading, gudtPlayer.udtPhysics.sngFacing + Pi
                Case (CANNON_FIXED Or CANNON_OCTA)
                    CreateBullet gudtPlayer.udtSystems.bytCannon, gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, -1, gudtPlayer.udtPhysics.sngSpeed, gudtPlayer.udtPhysics.sngHeading, gudtPlayer.udtPhysics.sngFacing
                    CreateBullet gudtPlayer.udtSystems.bytCannon, gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, -1, gudtPlayer.udtPhysics.sngSpeed, gudtPlayer.udtPhysics.sngHeading, gudtPlayer.udtPhysics.sngFacing + Pi / 4
                    CreateBullet gudtPlayer.udtSystems.bytCannon, gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, -1, gudtPlayer.udtPhysics.sngSpeed, gudtPlayer.udtPhysics.sngHeading, gudtPlayer.udtPhysics.sngFacing + Pi / 2
                    CreateBullet gudtPlayer.udtSystems.bytCannon, gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, -1, gudtPlayer.udtPhysics.sngSpeed, gudtPlayer.udtPhysics.sngHeading, gudtPlayer.udtPhysics.sngFacing + 3 * Pi / 4
                    CreateBullet gudtPlayer.udtSystems.bytCannon, gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, -1, gudtPlayer.udtPhysics.sngSpeed, gudtPlayer.udtPhysics.sngHeading, gudtPlayer.udtPhysics.sngFacing + Pi
                    CreateBullet gudtPlayer.udtSystems.bytCannon, gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, -1, gudtPlayer.udtPhysics.sngSpeed, gudtPlayer.udtPhysics.sngHeading, gudtPlayer.udtPhysics.sngFacing - 3 * Pi / 4
                    CreateBullet gudtPlayer.udtSystems.bytCannon, gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, -1, gudtPlayer.udtPhysics.sngSpeed, gudtPlayer.udtPhysics.sngHeading, gudtPlayer.udtPhysics.sngFacing - Pi / 2
                    CreateBullet gudtPlayer.udtSystems.bytCannon, gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, -1, gudtPlayer.udtPhysics.sngSpeed, gudtPlayer.udtPhysics.sngHeading, gudtPlayer.udtPhysics.sngFacing - Pi / 4
                'Single, turret
                Case CANNON_TURRET
                    'Single turret
                    With gudtPlayer
                        AccurateShot gudtObject(gudtRadar(.lngRadarObject).lngObject).udtPhysics.dblX, gudtObject(gudtRadar(.lngRadarObject).lngObject).udtPhysics.dblY, gudtObject(gudtRadar(.lngRadarObject).lngObject).udtPhysics.sngSpeed, gudtObject(gudtRadar(.lngRadarObject).lngObject).udtPhysics.sngHeading, .udtPhysics.dblX, .udtPhysics.dblY, .udtPhysics.sngSpeed, .udtPhysics.sngHeading, gudtCannon(.udtSystems.bytCannon).sngSpeed, sngHeading
                        CreateBullet .udtSystems.bytCannon, .udtPhysics.dblX, .udtPhysics.dblY, -1, .udtPhysics.sngSpeed, .udtPhysics.sngHeading, sngHeading
                    End With
                Case (CANNON_TURRET Or CANNON_DOUBLE)
                    'Double turret
                    With gudtPlayer
                        AccurateShot gudtObject(gudtRadar(.lngRadarObject).lngObject).udtPhysics.dblX, gudtObject(gudtRadar(.lngRadarObject).lngObject).udtPhysics.dblY, gudtObject(gudtRadar(.lngRadarObject).lngObject).udtPhysics.sngSpeed, gudtObject(gudtRadar(.lngRadarObject).lngObject).udtPhysics.sngHeading, .udtPhysics.dblX, .udtPhysics.dblY, .udtPhysics.sngSpeed, .udtPhysics.sngHeading, gudtCannon(.udtSystems.bytCannon).sngSpeed, sngHeading
                        PointOnLine .udtPhysics.dblX, .udtPhysics.dblY, sngHeading + Pi / 2, CANNON_SPREAD, dblX, dblY
                        CreateBullet .udtSystems.bytCannon, dblX, dblY, -1, .udtPhysics.sngSpeed, .udtPhysics.sngHeading, sngHeading
                        PointOnLine .udtPhysics.dblX, .udtPhysics.dblY, sngHeading - Pi / 2, CANNON_SPREAD, dblX, dblY
                        CreateBullet .udtSystems.bytCannon, dblX, dblY, -1, .udtPhysics.sngSpeed, .udtPhysics.sngHeading, sngHeading
                    End With
                Case (CANNON_TURRET Or CANNON_TRIPLE)
                    'Triple turret
                    With gudtPlayer
                        AccurateShot gudtObject(gudtRadar(.lngRadarObject).lngObject).udtPhysics.dblX, gudtObject(gudtRadar(.lngRadarObject).lngObject).udtPhysics.dblY, gudtObject(gudtRadar(.lngRadarObject).lngObject).udtPhysics.sngSpeed, gudtObject(gudtRadar(.lngRadarObject).lngObject).udtPhysics.sngHeading, .udtPhysics.dblX, .udtPhysics.dblY, .udtPhysics.sngSpeed, .udtPhysics.sngHeading, gudtCannon(.udtSystems.bytCannon).sngSpeed, sngHeading
                        CreateBullet .udtSystems.bytCannon, .udtPhysics.dblX, .udtPhysics.dblY, -1, .udtPhysics.sngSpeed, .udtPhysics.sngHeading, sngHeading
                        PointOnLine .udtPhysics.dblX, .udtPhysics.dblY, sngHeading + Pi / 2, CANNON_SPREAD * 1.5, dblX, dblY
                        CreateBullet .udtSystems.bytCannon, dblX, dblY, -1, .udtPhysics.sngSpeed, .udtPhysics.sngHeading, sngHeading
                        PointOnLine .udtPhysics.dblX, .udtPhysics.dblY, sngHeading - Pi / 2, CANNON_SPREAD * 1.5, dblX, dblY
                        CreateBullet .udtSystems.bytCannon, dblX, dblY, -1, .udtPhysics.sngSpeed, .udtPhysics.sngHeading, sngHeading
                    End With
                'Just a single, fixed forward cannon
                Case Else
                    CreateBullet gudtPlayer.udtSystems.bytCannon, gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, -1, gudtPlayer.udtPhysics.sngSpeed, gudtPlayer.udtPhysics.sngHeading, gudtPlayer.udtPhysics.sngFacing
            End Select
        End If
    End If

End Sub

Public Sub ObjectFireCannon(lngObject As Long)

Dim lngTemp As Long
Dim sngHeading As Single
Dim dblX As Double
Dim dblY As Double

    'Do we HAVE cannons?
    If gudtObject(lngObject).udtSystems.bytCannon <= 0 Then Exit Sub

    'Check crew..
    If gudtObject(lngObject).udtSystems.lngCrew <= 0 Then Exit Sub
    
    'Calculate duration of delay
    lngTemp = gudtCannon(gudtObject(lngObject).udtSystems.bytCannon).lngFireRate / (gudtObject(lngObject).udtSystems.lngCrew / gudtHull(gudtObject(lngObject).udtSystems.bytHull).lngMaxCrew)
    
    'Is it time?
    If gudtObject(lngObject).udtSystems.lngCannonLastFire + lngTemp <= glngGameTime Then
        'Do we have energy?
        If gudtObject(lngObject).udtSystems.sngEnergy >= gudtCannon(gudtObject(lngObject).udtSystems.bytCannon).sngInstantaneousConsumption Then
            'Remove energy
            gudtObject(lngObject).udtSystems.sngEnergy = gudtObject(lngObject).udtSystems.sngEnergy - gudtCannon(gudtObject(lngObject).udtSystems.bytCannon).sngInstantaneousConsumption
            'Store new lngCannonLastFire
            gudtObject(lngObject).udtSystems.lngCannonLastFire = glngGameTime
            'Fire
            Select Case gudtCannon(gudtObject(lngObject).udtSystems.bytCannon).lngCannonType
                Case (CANNON_FIXED Or CANNON_DOUBLE)
                    PointOnLine gudtObject(lngObject).udtPhysics.dblX, gudtObject(lngObject).udtPhysics.dblY, gudtObject(lngObject).udtPhysics.sngFacing + Pi / 2, CANNON_SPREAD, dblX, dblY
                    CreateBullet gudtObject(lngObject).udtSystems.bytCannon, dblX, dblY, lngObject, gudtObject(lngObject).udtPhysics.sngSpeed, gudtObject(lngObject).udtPhysics.sngHeading, gudtObject(lngObject).udtPhysics.sngFacing
                    PointOnLine gudtObject(lngObject).udtPhysics.dblX, gudtObject(lngObject).udtPhysics.dblY, gudtObject(lngObject).udtPhysics.sngFacing - Pi / 2, CANNON_SPREAD, dblX, dblY
                    CreateBullet gudtObject(lngObject).udtSystems.bytCannon, dblX, dblY, lngObject, gudtObject(lngObject).udtPhysics.sngSpeed, gudtObject(lngObject).udtPhysics.sngHeading, gudtObject(lngObject).udtPhysics.sngFacing
                Case (CANNON_FIXED Or CANNON_TRIPLE)
                    CreateBullet gudtObject(lngObject).udtSystems.bytCannon, gudtObject(lngObject).udtPhysics.dblX, gudtObject(lngObject).udtPhysics.dblY, lngObject, gudtObject(lngObject).udtPhysics.sngSpeed, gudtObject(lngObject).udtPhysics.sngHeading, gudtObject(lngObject).udtPhysics.sngFacing
                    CreateBullet gudtObject(lngObject).udtSystems.bytCannon, gudtObject(lngObject).udtPhysics.dblX, gudtObject(lngObject).udtPhysics.dblY, lngObject, gudtObject(lngObject).udtPhysics.sngSpeed, gudtObject(lngObject).udtPhysics.sngHeading, gudtObject(lngObject).udtPhysics.sngFacing + Pi / 9
                    CreateBullet gudtObject(lngObject).udtSystems.bytCannon, gudtObject(lngObject).udtPhysics.dblX, gudtObject(lngObject).udtPhysics.dblY, lngObject, gudtObject(lngObject).udtPhysics.sngSpeed, gudtObject(lngObject).udtPhysics.sngHeading, gudtObject(lngObject).udtPhysics.sngFacing - Pi / 9
                Case (CANNON_FIXED Or CANNON_QUAD)
                    CreateBullet gudtObject(lngObject).udtSystems.bytCannon, gudtObject(lngObject).udtPhysics.dblX, gudtObject(lngObject).udtPhysics.dblY, lngObject, gudtObject(lngObject).udtPhysics.sngSpeed, gudtObject(lngObject).udtPhysics.sngHeading, gudtObject(lngObject).udtPhysics.sngFacing
                    CreateBullet gudtObject(lngObject).udtSystems.bytCannon, gudtObject(lngObject).udtPhysics.dblX, gudtObject(lngObject).udtPhysics.dblY, lngObject, gudtObject(lngObject).udtPhysics.sngSpeed, gudtObject(lngObject).udtPhysics.sngHeading, gudtObject(lngObject).udtPhysics.sngFacing + Pi / 2
                    CreateBullet gudtObject(lngObject).udtSystems.bytCannon, gudtObject(lngObject).udtPhysics.dblX, gudtObject(lngObject).udtPhysics.dblY, lngObject, gudtObject(lngObject).udtPhysics.sngSpeed, gudtObject(lngObject).udtPhysics.sngHeading, gudtObject(lngObject).udtPhysics.sngFacing - Pi / 2
                    CreateBullet gudtObject(lngObject).udtSystems.bytCannon, gudtObject(lngObject).udtPhysics.dblX, gudtObject(lngObject).udtPhysics.dblY, lngObject, gudtObject(lngObject).udtPhysics.sngSpeed, gudtObject(lngObject).udtPhysics.sngHeading, gudtObject(lngObject).udtPhysics.sngFacing + Pi
                Case (CANNON_FIXED Or CANNON_OCTA)
                    CreateBullet gudtObject(lngObject).udtSystems.bytCannon, gudtObject(lngObject).udtPhysics.dblX, gudtObject(lngObject).udtPhysics.dblY, lngObject, gudtObject(lngObject).udtPhysics.sngSpeed, gudtObject(lngObject).udtPhysics.sngHeading, gudtObject(lngObject).udtPhysics.sngFacing
                    CreateBullet gudtObject(lngObject).udtSystems.bytCannon, gudtObject(lngObject).udtPhysics.dblX, gudtObject(lngObject).udtPhysics.dblY, lngObject, gudtObject(lngObject).udtPhysics.sngSpeed, gudtObject(lngObject).udtPhysics.sngHeading, gudtObject(lngObject).udtPhysics.sngFacing + Pi / 4
                    CreateBullet gudtObject(lngObject).udtSystems.bytCannon, gudtObject(lngObject).udtPhysics.dblX, gudtObject(lngObject).udtPhysics.dblY, lngObject, gudtObject(lngObject).udtPhysics.sngSpeed, gudtObject(lngObject).udtPhysics.sngHeading, gudtObject(lngObject).udtPhysics.sngFacing + Pi / 2
                    CreateBullet gudtObject(lngObject).udtSystems.bytCannon, gudtObject(lngObject).udtPhysics.dblX, gudtObject(lngObject).udtPhysics.dblY, lngObject, gudtObject(lngObject).udtPhysics.sngSpeed, gudtObject(lngObject).udtPhysics.sngHeading, gudtObject(lngObject).udtPhysics.sngFacing + 3 * Pi / 4
                    CreateBullet gudtObject(lngObject).udtSystems.bytCannon, gudtObject(lngObject).udtPhysics.dblX, gudtObject(lngObject).udtPhysics.dblY, lngObject, gudtObject(lngObject).udtPhysics.sngSpeed, gudtObject(lngObject).udtPhysics.sngHeading, gudtObject(lngObject).udtPhysics.sngFacing + Pi
                    CreateBullet gudtObject(lngObject).udtSystems.bytCannon, gudtObject(lngObject).udtPhysics.dblX, gudtObject(lngObject).udtPhysics.dblY, lngObject, gudtObject(lngObject).udtPhysics.sngSpeed, gudtObject(lngObject).udtPhysics.sngHeading, gudtObject(lngObject).udtPhysics.sngFacing - 3 * Pi / 4
                    CreateBullet gudtObject(lngObject).udtSystems.bytCannon, gudtObject(lngObject).udtPhysics.dblX, gudtObject(lngObject).udtPhysics.dblY, lngObject, gudtObject(lngObject).udtPhysics.sngSpeed, gudtObject(lngObject).udtPhysics.sngHeading, gudtObject(lngObject).udtPhysics.sngFacing - Pi / 2
                    CreateBullet gudtObject(lngObject).udtSystems.bytCannon, gudtObject(lngObject).udtPhysics.dblX, gudtObject(lngObject).udtPhysics.dblY, lngObject, gudtObject(lngObject).udtPhysics.sngSpeed, gudtObject(lngObject).udtPhysics.sngHeading, gudtObject(lngObject).udtPhysics.sngFacing - Pi / 4
                'Single, turret
                Case CANNON_TURRET
                    'Single turret
                    With gudtObject(lngObject)
                        If .udtAI.lngTarget = TARGET_PLAYER Then
                            AccurateShot gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, gudtPlayer.udtPhysics.sngSpeed, gudtPlayer.udtPhysics.sngHeading, .udtPhysics.dblX, .udtPhysics.dblY, .udtPhysics.sngSpeed, .udtPhysics.sngHeading, gudtCannon(.udtSystems.bytCannon).sngSpeed, sngHeading
                        Else
                            AccurateShot gudtObject(.udtAI.lngTarget).udtPhysics.dblX, gudtObject(.udtAI.lngTarget).udtPhysics.dblY, gudtObject(.udtAI.lngTarget).udtPhysics.sngSpeed, gudtObject(.udtAI.lngTarget).udtPhysics.sngHeading, .udtPhysics.dblX, .udtPhysics.dblY, .udtPhysics.sngSpeed, .udtPhysics.sngHeading, gudtCannon(.udtSystems.bytCannon).sngSpeed, sngHeading
                        End If
                        CreateBullet .udtSystems.bytCannon, .udtPhysics.dblX, .udtPhysics.dblY, lngObject, .udtPhysics.sngSpeed, .udtPhysics.sngHeading, sngHeading
                    End With
                Case (CANNON_TURRET Or CANNON_DOUBLE)
                    'Double turret
                    With gudtObject(lngObject)
                        If .udtAI.lngTarget = TARGET_PLAYER Then
                            AccurateShot gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, gudtPlayer.udtPhysics.sngSpeed, gudtPlayer.udtPhysics.sngHeading, .udtPhysics.dblX, .udtPhysics.dblY, .udtPhysics.sngSpeed, .udtPhysics.sngHeading, gudtCannon(.udtSystems.bytCannon).sngSpeed, sngHeading
                        Else
                            AccurateShot gudtObject(.udtAI.lngTarget).udtPhysics.dblX, gudtObject(.udtAI.lngTarget).udtPhysics.dblY, gudtObject(.udtAI.lngTarget).udtPhysics.sngSpeed, gudtObject(.udtAI.lngTarget).udtPhysics.sngHeading, .udtPhysics.dblX, .udtPhysics.dblY, .udtPhysics.sngSpeed, .udtPhysics.sngHeading, gudtCannon(.udtSystems.bytCannon).sngSpeed, sngHeading
                        End If
                        PointOnLine .udtPhysics.dblX, .udtPhysics.dblY, sngHeading + Pi / 2, CANNON_SPREAD, dblX, dblY
                        CreateBullet .udtSystems.bytCannon, dblX, dblY, lngObject, .udtPhysics.sngSpeed, .udtPhysics.sngHeading, sngHeading
                        PointOnLine .udtPhysics.dblX, .udtPhysics.dblY, sngHeading - Pi / 2, CANNON_SPREAD, dblX, dblY
                        CreateBullet .udtSystems.bytCannon, dblX, dblY, lngObject, .udtPhysics.sngSpeed, .udtPhysics.sngHeading, sngHeading
                    End With
                Case (CANNON_TURRET Or CANNON_TRIPLE)
                    'Triple turret
                    With gudtObject(lngObject)
                        If .udtAI.lngTarget = TARGET_PLAYER Then
                            AccurateShot gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, gudtPlayer.udtPhysics.sngSpeed, gudtPlayer.udtPhysics.sngHeading, .udtPhysics.dblX, .udtPhysics.dblY, .udtPhysics.sngSpeed, .udtPhysics.sngHeading, gudtCannon(.udtSystems.bytCannon).sngSpeed, sngHeading
                        Else
                            AccurateShot gudtObject(.udtAI.lngTarget).udtPhysics.dblX, gudtObject(.udtAI.lngTarget).udtPhysics.dblY, gudtObject(.udtAI.lngTarget).udtPhysics.sngSpeed, gudtObject(.udtAI.lngTarget).udtPhysics.sngHeading, .udtPhysics.dblX, .udtPhysics.dblY, .udtPhysics.sngSpeed, .udtPhysics.sngHeading, gudtCannon(.udtSystems.bytCannon).sngSpeed, sngHeading
                        End If
                        CreateBullet .udtSystems.bytCannon, .udtPhysics.dblX, .udtPhysics.dblY, lngObject, .udtPhysics.sngSpeed, .udtPhysics.sngHeading, sngHeading
                        PointOnLine .udtPhysics.dblX, .udtPhysics.dblY, sngHeading + Pi / 2, CANNON_SPREAD * 1.5, dblX, dblY
                        CreateBullet .udtSystems.bytCannon, dblX, dblY, lngObject, .udtPhysics.sngSpeed, .udtPhysics.sngHeading, sngHeading
                        PointOnLine .udtPhysics.dblX, .udtPhysics.dblY, sngHeading - Pi / 2, CANNON_SPREAD * 1.5, dblX, dblY
                        CreateBullet .udtSystems.bytCannon, dblX, dblY, lngObject, .udtPhysics.sngSpeed, .udtPhysics.sngHeading, sngHeading
                    End With
                'Just a single, fixed forward cannon
                Case Else
                    CreateBullet gudtObject(lngObject).udtSystems.bytCannon, gudtObject(lngObject).udtPhysics.dblX, gudtObject(lngObject).udtPhysics.dblY, lngObject, gudtObject(lngObject).udtPhysics.sngSpeed, gudtObject(lngObject).udtPhysics.sngHeading, gudtObject(lngObject).udtPhysics.sngFacing
            End Select
        End If
    End If

End Sub

Sub CheckCannonCollisions()

Dim i As Long
Dim j As Long
Dim blnCheck As Boolean
Dim sngShieldPercent As Single

    'Check collisions with player..
    i = 0
    Do While i <= glngNumBullets - 1
        'Ensure the player isn't the owner
        If gudtBullet(i).lngOwner <> -1 Then
            'Check distance
            If GetDist(gudtBullet(i).dblX, gudtBullet(i).dblY, gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY) <= gudtPlayer.udtSprite.intWidth \ 2 + BULLET_WIDTH \ 2 Then
                'Calc shield percent
                sngShieldPercent = gudtPlayer.udtSystems.sngShieldEnergy / gudtShield(gudtPlayer.udtSystems.bytShield).lngMaxEnergy
                'We have contact! Check shields
                If gudtPlayer.udtSystems.sngShieldEnergy > 0 Then
                    'Shields are up, display them
                    gudtPlayer.udtInfo.blnShieldUp = True
                    gudtPlayer.udtInfo.lngShieldDown = glngGameTime + SHIELD_DURATION
                End If
                'Display an explosion?
                If sngShieldPercent < 1 Then
                    'Shields are not 100%, display explosion
                    CreateExplosion gudtBullet(i).dblX, gudtBullet(i).dblY, gudtPlayer.udtPhysics.sngSpeed, gudtPlayer.udtPhysics.sngHeading
                End If
                'Calc energy loss
                If gudtPlayer.udtSystems.sngEnergy >= gudtCannon(gudtBullet(i).bytCannon).sngConcussiveDamage * sngShieldPercent Then
                    'Take this straight from sngEnergy
                    gudtPlayer.udtSystems.sngEnergy = gudtPlayer.udtSystems.sngEnergy - gudtCannon(gudtBullet(i).bytCannon).sngConcussiveDamage * sngShieldPercent
                Else
                    'Take some from shields
                    gudtPlayer.udtSystems.sngShieldEnergy = (gudtPlayer.udtSystems.sngEnergy + gudtPlayer.udtSystems.sngShieldEnergy) - gudtCannon(gudtBullet(i).bytCannon).sngConcussiveDamage * sngShieldPercent
                End If
                'Calc damage
                gudtPlayer.udtSystems.lngArmour = gudtPlayer.udtSystems.lngArmour - gudtCannon(gudtBullet(i).bytCannon).sngConcussiveDamage * (1 - sngShieldPercent)
                'Calc crew loss
                gudtPlayer.udtSystems.lngCrew = gudtPlayer.udtSystems.lngCrew - gudtCannon(gudtBullet(i).bytCannon).sngRadiationDamage * (1 - sngShieldPercent)
                If gudtPlayer.udtSystems.lngCrew < 0 Then gudtPlayer.udtSystems.lngCrew = 0
                'Dead?
                If gudtPlayer.udtSystems.lngArmour <= 0 Then PlayerDead
                'Remove the bullet
                DeleteBullet i
                i = i - 1
            End If
        End If
        'Increment
        i = i + 1
    Loop

    'Check collisions with objects
    i = 0
    Do While i <= glngNumBullets - 1
        'Loop through all objects.. (ugh!)
        For j = 0 To UBound(gudtObject)
            'Ensure this isn't the owner, nonexistant, or a planet/star
            If (gudtBullet(i).lngOwner <> j) And (gudtObject(j).udtInfo.blnPlanet <> True) And (gudtObject(j).udtInfo.blnStar <> True) And (gudtObject(j).blnExists = True) Then
                'Is this bullet owned by the player?
                blnCheck = False
                If gudtBullet(i).lngOwner = -1 Then
                    blnCheck = True
                'Ensure this isn't in the same race
                ElseIf gudtObject(gudtBullet(i).lngOwner).udtInfo.bytRace <> gudtObject(j).udtInfo.bytRace Then
                    blnCheck = True
                End If
                'Check this one?
                If blnCheck = True Then
                    'Check distance
                    If GetDist(gudtObject(j).udtPhysics.dblX, gudtObject(j).udtPhysics.dblY, gudtBullet(i).dblX, gudtBullet(i).dblY) <= gudtObject(j).udtSprite.intWidth \ 2 + BULLET_WIDTH \ 2 Then
                        'We have contact!
                        'Calc shield percent
                        sngShieldPercent = 2 * gudtObject(j).udtSystems.sngEnergy / gudtGenerator(gudtObject(j).udtSystems.bytGenerator).lngMaxBattery
                        If sngShieldPercent > 1 Then sngShieldPercent = 1
                        'Do we have any shield to display?
                        If sngShieldPercent > 0 Then
                            gudtObject(j).udtInfo.blnShieldUp = True
                            gudtObject(j).udtInfo.lngShieldDown = glngGameTime + SHIELD_DURATION
                        End If
                        'If energy is less than 50%, show explosion (presumably the shields would be weakened)
                        If sngShieldPercent < 1 Then
                            CreateExplosion gudtBullet(i).dblX, gudtBullet(i).dblY, gudtObject(j).udtPhysics.sngSpeed, gudtObject(j).udtPhysics.sngHeading
                        End If
                        'Calc energy loss
                        gudtObject(j).udtSystems.sngEnergy = gudtObject(j).udtSystems.sngEnergy - gudtCannon(gudtBullet(i).bytCannon).sngConcussiveDamage * sngShieldPercent
                        'Calc damage
                        gudtObject(j).udtSystems.lngArmour = gudtObject(j).udtSystems.lngArmour - gudtCannon(gudtBullet(i).bytCannon).sngConcussiveDamage * (1 - sngShieldPercent)
                        'Calc crew loss
                        gudtObject(j).udtSystems.lngCrew = gudtObject(j).udtSystems.lngCrew - gudtCannon(gudtBullet(i).bytCannon).sngRadiationDamage * (1 - sngShieldPercent)
                        If gudtObject(j).udtSystems.lngCrew < 0 Then gudtObject(j).udtSystems.lngCrew = 0
                        'Dead?
                        If gudtObject(j).udtSystems.lngArmour <= 0 Then
                            ObjectDead j
                            'Current radar object?
                            If gudtPlayer.lngRadarObject >= 0 Then
                                If gudtRadar(gudtPlayer.lngRadarObject).lngObject = j Then
                                    'Tab to another
                                    Tactical.RadarTab
                                    Tactical.RadarEnemyTab
                                End If
                            End If
                        End If
                        'Remove the bullet
                        DeleteBullet i
                        i = i - 1
                        Exit For
                    End If
                End If
            End If
        Next j
        'Increment
        i = i + 1
    Loop

End Sub

Sub CheckMissileCollisions()

Dim i As Long
Dim j As Long
Dim blnCheck As Boolean
Dim sngShieldPercent As Single
Dim sngEnergyLoss As Single

    'Check collisions with player..
    i = 0
    Do While i <= glngNumLiveMissiles - 1
        'Ensure the player isn't the owner
        If gudtLiveMissile(i).lngOwner <> -1 Then
            'Check distance
            If GetDist(gudtLiveMissile(i).dblX, gudtLiveMissile(i).dblY, gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY) <= gudtPlayer.udtSprite.intWidth \ 2 + gudtMissile(gudtLiveMissile(i).bytMissile).udtSprite.intWidth Then
                'Calc shield percent
                sngShieldPercent = gudtPlayer.udtSystems.sngShieldEnergy / gudtShield(gudtPlayer.udtSystems.bytShield).lngMaxEnergy
                'We have contact! Check shields
                If gudtPlayer.udtSystems.sngShieldEnergy > 0 Then
                    'Shields are up, display them
                    gudtPlayer.udtInfo.blnShieldUp = True
                    gudtPlayer.udtInfo.lngShieldDown = glngGameTime + SHIELD_DURATION
                End If
                'Display an explosion
                CreateExplosion gudtLiveMissile(i).dblX, gudtLiveMissile(i).dblY, gudtPlayer.udtPhysics.sngSpeed, gudtPlayer.udtPhysics.sngHeading
                'Calc energy loss
                If gudtPlayer.udtSystems.sngEnergy >= gudtMissile(gudtLiveMissile(i).bytMissile).sngConcussiveDamage * sngShieldPercent Then
                    'Take this straight from sngEnergy
                    gudtPlayer.udtSystems.sngEnergy = gudtPlayer.udtSystems.sngEnergy - gudtMissile(gudtLiveMissile(i).bytMissile).sngConcussiveDamage * sngShieldPercent
                Else
                    'Take some from shields
                    gudtPlayer.udtSystems.sngShieldEnergy = (gudtPlayer.udtSystems.sngEnergy + gudtPlayer.udtSystems.sngShieldEnergy) - gudtMissile(gudtLiveMissile(i).bytMissile).sngConcussiveDamage * sngShieldPercent
                End If
                'Calc damage
                gudtPlayer.udtSystems.lngArmour = gudtPlayer.udtSystems.lngArmour - gudtMissile(gudtLiveMissile(i).bytMissile).sngConcussiveDamage * (1 - sngShieldPercent)
                'Calc crew loss
                gudtPlayer.udtSystems.lngCrew = gudtPlayer.udtSystems.lngCrew - gudtMissile(gudtLiveMissile(i).bytMissile).sngRadiationDamage * (1 - sngShieldPercent)
                If gudtPlayer.udtSystems.lngCrew < 0 Then gudtPlayer.udtSystems.lngCrew = 0
                'Dead?
                If gudtPlayer.udtSystems.lngArmour <= 0 Then PlayerDead
                'Remove the missile
                DeleteMissile i
                i = i - 1
            End If
        End If
        'Increment
        i = i + 1
    Loop

    'Check collisions with objects
    i = 0
    Do While i <= glngNumLiveMissiles - 1
        'Loop through all objects.. (ugh!)
        For j = 0 To UBound(gudtObject)
            'Ensure this isn't the owner, nonexistant, or a planet/star
            If (gudtLiveMissile(i).lngOwner <> j) And (gudtObject(j).udtInfo.blnPlanet <> True) And (gudtObject(j).udtInfo.blnStar <> True) And (gudtObject(j).blnExists = True) Then
                'Is this missile owned by the player?
                blnCheck = False
                If gudtLiveMissile(i).lngOwner = -1 Then
                    blnCheck = True
                'Ensure this isn't in the same race
                ElseIf gudtObject(gudtLiveMissile(i).lngOwner).udtInfo.bytRace <> gudtObject(j).udtInfo.bytRace Then
                    blnCheck = True
                End If
                'Check this one?
                If blnCheck = True Then
                    'Check distance
                    If GetDist(gudtObject(j).udtPhysics.dblX, gudtObject(j).udtPhysics.dblY, gudtLiveMissile(i).dblX, gudtLiveMissile(i).dblY) <= gudtObject(j).udtSprite.intWidth \ 2 + gudtMissile(gudtLiveMissile(i).bytMissile).udtSprite.intWidth Then
                        'We have contact!
                        'Calc shield percent
                        sngShieldPercent = 2 * gudtObject(j).udtSystems.sngEnergy / gudtGenerator(gudtObject(j).udtSystems.bytGenerator).lngMaxBattery
                        If sngShieldPercent > 1 Then sngShieldPercent = 1
                        'Do we have any shield to display?
                        If sngShieldPercent > 0 Then
                            gudtObject(j).udtInfo.blnShieldUp = True
                            gudtObject(j).udtInfo.lngShieldDown = glngGameTime + SHIELD_DURATION
                        End If
                        'Show explosion
                        CreateExplosion gudtLiveMissile(i).dblX, gudtLiveMissile(i).dblY, gudtObject(j).udtPhysics.sngSpeed, gudtObject(j).udtPhysics.sngHeading
                        'Calc energy loss
                        sngEnergyLoss = gudtObject(j).udtSystems.sngEnergy
                        gudtObject(j).udtSystems.sngEnergy = gudtObject(j).udtSystems.sngEnergy - gudtMissile(gudtLiveMissile(i).bytMissile).sngConcussiveDamage * sngShieldPercent
                        'Calc damage
                        'Are ALL the shields gone?
                        If gudtObject(j).udtSystems.sngEnergy <= 0 Then
                            'If so, transfer whatever wasn't absorbed by shields to the armour
                            gudtObject(j).udtSystems.lngArmour = gudtObject(j).udtSystems.lngArmour - (gudtMissile(gudtLiveMissile(i).bytMissile).sngConcussiveDamage - sngEnergyLoss)
                        Else
                            gudtObject(j).udtSystems.lngArmour = gudtObject(j).udtSystems.lngArmour - gudtMissile(gudtLiveMissile(i).bytMissile).sngConcussiveDamage * (1 - sngShieldPercent)
                        End If
                        'Calc crew loss
                        gudtObject(j).udtSystems.lngCrew = gudtObject(j).udtSystems.lngCrew - gudtMissile(gudtLiveMissile(i).bytMissile).sngRadiationDamage * (1 - sngShieldPercent)
                        If gudtObject(j).udtSystems.lngCrew < 0 Then gudtObject(j).udtSystems.lngCrew = 0
                        'Dead?
                        If gudtObject(j).udtSystems.lngArmour <= 0 Then
                            ObjectDead j
                            'Current radar object?
                            If gudtPlayer.lngRadarObject >= 0 Then
                                If gudtRadar(gudtPlayer.lngRadarObject).lngObject = j Then
                                    'Tab to another
                                    Tactical.RadarTab
                                    Tactical.RadarEnemyTab
                                End If
                            End If
                        End If
                        'Remove the missile
                        DeleteMissile i
                        i = i - 1
                        Exit For
                    End If
                End If
            End If
        Next j
        'Increment
        i = i + 1
    Loop

End Sub

Sub PlayerDead()

Dim i As Long

    'Dead dead dead
    gblnPlayerDead = True
    gblnPlayerExploded = False
    gblnPlayerDeadMessage = False
    glngPlayerExplodingStart = glngGameTime
    glngPlayerExplosionNum = 0
    
    'Stop stuff
    gudtPlayer.udtAI.bytAction = AI_NONE
    If gudtPlayer.udtAI.blnLaserFire = True Then
        gudtPlayer.udtAI.bytAction = ACTION_NONE
        gudtPlayer.udtAI.blnLaserFire = False
        DSound.StopSound gudtPlayer.udtAI.lngLaserSound
    End If
    If gudtPlayer.udtAI.blnThrusting = True Then
        gudtPlayer.udtAI.blnThrusting = False
        DSound.StopSound gudtPlayer.udtAI.lngThrustSound
    End If
    
    'Stop objects from attacking player!
    For i = 0 To UBound(gudtObject)
        'Is target player?
        If gudtObject(i).udtAI.lngTarget = -1 Then
            gudtObject(i).udtAI.lngTarget = 0
            gudtObject(i).udtAI.bytAction = AI_NONE
        End If
    Next i

End Sub

Sub ObjectDead(lngIndex As Long)

Dim i As Long

    'You no longer exist!
    gudtObject(lngIndex).blnExists = False
    
    For i = 0 To UBound(gudtObject)
        'Is this the target?
        If gudtObject(i).udtAI.lngTarget = lngIndex Then
            gudtObject(i).udtAI.lngTarget = TARGET_NONE
        End If
    Next i

    'Is this a missile target?
    For i = 0 To glngNumLiveMissiles - 1
        If gudtLiveMissile(i).lngTarget = lngIndex Then gudtLiveMissile(i).lngTarget = TARGET_NONE
    Next i

    'Is this the player's target?
    If gudtPlayer.udtAI.lngTarget = lngIndex Then
        gudtPlayer.udtAI.lngTarget = TARGET_NONE
        Tactical.RadarTab
        Tactical.RadarEnemyTab
    End If

End Sub

Function FindShield(intWidth As Integer) As Byte

    'Return the appropriate shield constant
    Select Case intWidth
        Case 10
            FindShield = SHIELD_0
        Case 20
            FindShield = SHIELD_1
        Case 40
            FindShield = SHIELD_2
        Case 60
            FindShield = SHIELD_3
        Case 80
            FindShield = SHIELD_4
        Case 100
            FindShield = SHIELD_5
        Case Else
            FindShield = SHIELD_9
    End Select

End Function

Sub CreateLaser(lngOwner As Long, lngTarget As Long, bytType As Byte)

    'Make a new spot
    ReDim Preserve gudtLaserDisplay(glngNumLaserDisplay)
    glngNumLaserDisplay = glngNumLaserDisplay + 1
    
    'Enter the data
    gudtLaserDisplay(glngNumLaserDisplay - 1).bytType = bytType
    gudtLaserDisplay(glngNumLaserDisplay - 1).lngOwner = lngOwner
    gudtLaserDisplay(glngNumLaserDisplay - 1).lngTarget = lngTarget

End Sub

Sub PlayerFireLaser()

Dim sngEnergy As Single
Dim sngConcussive As Single
Dim sngRadiation As Single
Dim sngShieldPercent As Single

    'We're not firing the laser yet!
    gudtPlayer.udtAI.blnLaserFire = False

    'If there's no radarobject, then there's no laser!
    If gudtPlayer.lngRadarObject < 0 Then Exit Sub

    'Is the object a star or planet?
    If (gudtObject(gudtRadar(gudtPlayer.lngRadarObject).lngObject).udtInfo.blnPlanet = True) Or (gudtObject(gudtRadar(gudtPlayer.lngRadarObject).lngObject).udtInfo.blnStar = True) Then Exit Sub

    'Check crew..
    If gudtPlayer.udtSystems.lngCrew <= 0 Then Exit Sub
    
    'Check weapons energy
    If gudtPlayer.udtSystems.sngWeaponEnergy <= 0 Then Exit Sub

    'Does the player have lasers?
    If gudtPlayer.udtSystems.bytLaser = 0 Then Exit Sub
    
    'Check the distance
    If gudtObject(gudtRadar(gudtPlayer.lngRadarObject).lngObject).udtInfo.dblDistance > gudtLaser(gudtPlayer.udtSystems.bytLaser).lngRange Then Exit Sub
    
    'Take the energy..
    sngEnergy = gudtLaser(gudtPlayer.udtSystems.bytLaser).sngFireConsumption * glngElapsed * (gudtPlayer.udtSystems.lngCrew / gudtHull(gudtPlayer.udtSystems.bytHull).lngMaxCrew) * (gudtPlayer.udtSystems.sngWeaponEnergy / (gudtCannon(gudtPlayer.udtSystems.bytCannon).lngMaxEnergy + gudtLaser(gudtPlayer.udtSystems.bytLaser).lngMaxEnergy))
    If gudtPlayer.udtSystems.sngEnergy >= sngEnergy Then
        gudtPlayer.udtSystems.sngEnergy = gudtPlayer.udtSystems.sngEnergy - sngEnergy
    Else
        gudtPlayer.udtSystems.sngWeaponEnergy = gudtPlayer.udtSystems.sngWeaponEnergy + gudtPlayer.udtSystems.sngEnergy - sngEnergy
    End If

    'Calc the damage (dependent upon weapon energy and crew)
    sngConcussive = (gudtLaser(gudtPlayer.udtSystems.bytLaser).sngConcussiveDamage * glngElapsed) * (gudtPlayer.udtSystems.lngCrew / gudtHull(gudtPlayer.udtSystems.bytHull).lngMaxCrew) * (gudtPlayer.udtSystems.sngWeaponEnergy / (gudtCannon(gudtPlayer.udtSystems.bytCannon).lngMaxEnergy + gudtLaser(gudtPlayer.udtSystems.bytLaser).lngMaxEnergy))
    sngRadiation = (gudtLaser(gudtPlayer.udtSystems.bytLaser).sngRadiationDamage * glngElapsed) * (gudtPlayer.udtSystems.lngCrew / gudtHull(gudtPlayer.udtSystems.bytHull).lngMaxCrew) * (gudtPlayer.udtSystems.sngWeaponEnergy / (gudtCannon(gudtPlayer.udtSystems.bytCannon).lngMaxEnergy + gudtLaser(gudtPlayer.udtSystems.bytLaser).lngMaxEnergy))
    
    'Calc the shield percent
    sngShieldPercent = 2 * gudtObject(gudtRadar(gudtPlayer.lngRadarObject).lngObject).udtSystems.sngEnergy / gudtGenerator(gudtObject(gudtRadar(gudtPlayer.lngRadarObject).lngObject).udtSystems.bytGenerator).lngMaxBattery
    If sngShieldPercent > 1 Then sngShieldPercent = 1
    'Do we have any shield to display?
    If sngShieldPercent > 0 Then
        gudtObject(gudtRadar(gudtPlayer.lngRadarObject).lngObject).udtInfo.blnShieldUp = True
        gudtObject(gudtRadar(gudtPlayer.lngRadarObject).lngObject).udtInfo.lngShieldDown = glngGameTime + SHIELD_DURATION
    End If
    'Calc energy loss
    gudtObject(gudtRadar(gudtPlayer.lngRadarObject).lngObject).udtSystems.sngEnergy = gudtObject(gudtRadar(gudtPlayer.lngRadarObject).lngObject).udtSystems.sngEnergy - sngConcussive * sngShieldPercent
    'Calc damage
    gudtObject(gudtRadar(gudtPlayer.lngRadarObject).lngObject).udtSystems.lngArmour = gudtObject(gudtRadar(gudtPlayer.lngRadarObject).lngObject).udtSystems.lngArmour - sngConcussive * (1 - sngShieldPercent)
    'Calc crew loss
    gudtObject(gudtRadar(gudtPlayer.lngRadarObject).lngObject).udtSystems.lngCrew = gudtObject(gudtRadar(gudtPlayer.lngRadarObject).lngObject).udtSystems.lngCrew - sngRadiation * (1 - sngShieldPercent)
    If gudtObject(gudtRadar(gudtPlayer.lngRadarObject).lngObject).udtSystems.lngCrew < 0 Then gudtObject(gudtRadar(gudtPlayer.lngRadarObject).lngObject).udtSystems.lngCrew = 0
    'Dead?
    If gudtObject(gudtRadar(gudtPlayer.lngRadarObject).lngObject).udtSystems.lngArmour <= 0 Then
        'Make an explosion
        CreateExplosion gudtObject(gudtRadar(gudtPlayer.lngRadarObject).lngObject).udtPhysics.dblX, gudtObject(gudtRadar(gudtPlayer.lngRadarObject).lngObject).udtPhysics.dblY, gudtObject(gudtRadar(gudtPlayer.lngRadarObject).lngObject).udtPhysics.sngSpeed, gudtObject(gudtRadar(gudtPlayer.lngRadarObject).lngObject).udtPhysics.sngHeading
        'It's dead
        ObjectDead gudtRadar(gudtPlayer.lngRadarObject).lngObject
        'Current radar object?
        If gudtRadar(gudtPlayer.lngRadarObject).lngObject = gudtRadar(gudtPlayer.lngRadarObject).lngObject Then
            'Tab to another
            Tactical.RadarTab
            Tactical.RadarEnemyTab
        End If
    Else
        'Create a new laser
        CreateLaser -1, gudtRadar(gudtPlayer.lngRadarObject).lngObject, gudtPlayer.udtSystems.bytLaser
        'We are indeed firing the laser
        gudtPlayer.udtAI.blnLaserFire = True
    End If

End Sub

Sub ObjectFireLaser(lngObject As Long)

Dim sngEnergy As Single
Dim sngConcussive As Single
Dim sngRadiation As Single
Dim sngShieldPercent As Single
Dim sngEnergyLoss As Single
Dim sngArmourLoss As Single
Dim sngCrewLoss As Single
Dim blnPlayer As Boolean
Dim blnDead As Boolean

    'Lasers OFF!
    gudtObject(lngObject).udtAI.blnLaserFire = False

    'If there's no target, then there's no laser!
    If gudtObject(lngObject).udtAI.lngTarget = TARGET_NONE Then Exit Sub

    'Is the object a star or planet?
    If gudtObject(lngObject).udtAI.lngTarget <> TARGET_PLAYER Then If (gudtObject(gudtObject(lngObject).udtAI.lngTarget).udtInfo.blnPlanet = True) Or (gudtObject(gudtObject(lngObject).udtAI.lngTarget).udtInfo.blnStar = True) Then Exit Sub

    'Check crew..
    If gudtObject(lngObject).udtSystems.lngCrew <= 0 Then Exit Sub
    
    'Check weapons energy
    If gudtObject(lngObject).udtSystems.sngEnergy <= 0 Then Exit Sub

    'Does the object have lasers?
    If gudtObject(lngObject).udtSystems.bytLaser = 0 Then Exit Sub
    
    'Player?
    If gudtObject(lngObject).udtAI.lngTarget = TARGET_PLAYER Then
        blnPlayer = True
    Else
        blnPlayer = False
    End If
    
    'Is the target the player?
    If blnPlayer = True Then
        'Check the distance
        If GetDist(gudtObject(lngObject).udtPhysics.dblX, gudtObject(lngObject).udtPhysics.dblY, gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY) > gudtLaser(gudtObject(lngObject).udtSystems.bytLaser).lngRange Then Exit Sub
    Else
        'Check the distance
        If GetDist(gudtObject(lngObject).udtPhysics.dblX, gudtObject(lngObject).udtPhysics.dblY, gudtObject(gudtObject(lngObject).udtAI.lngTarget).udtPhysics.dblX, gudtObject(gudtObject(lngObject).udtAI.lngTarget).udtPhysics.dblY) > gudtLaser(gudtObject(lngObject).udtSystems.bytLaser).lngRange Then Exit Sub
    End If
    
    'Take the energy..
    sngEnergy = gudtLaser(gudtObject(lngObject).udtSystems.bytLaser).sngFireConsumption * glngElapsed * (gudtObject(lngObject).udtSystems.lngCrew / gudtHull(gudtObject(lngObject).udtSystems.bytHull).lngMaxCrew) * (gudtObject(lngObject).udtSystems.sngEnergy / (gudtGenerator(gudtObject(lngObject).udtSystems.bytGenerator).lngMaxBattery))
    gudtObject(lngObject).udtSystems.sngEnergy = gudtObject(lngObject).udtSystems.sngEnergy - sngEnergy

    'Calc the damage (dependent upon weapon energy and crew)
    sngConcussive = (gudtLaser(gudtObject(lngObject).udtSystems.bytLaser).sngConcussiveDamage * glngElapsed) * (gudtObject(lngObject).udtSystems.lngCrew / gudtHull(gudtObject(lngObject).udtSystems.bytHull).lngMaxCrew) * (gudtObject(lngObject).udtSystems.sngEnergy / gudtGenerator(gudtObject(lngObject).udtSystems.bytGenerator).lngMaxBattery)
    sngRadiation = (gudtLaser(gudtObject(lngObject).udtSystems.bytLaser).sngRadiationDamage * glngElapsed) * (gudtObject(lngObject).udtSystems.lngCrew / gudtHull(gudtObject(lngObject).udtSystems.bytHull).lngMaxCrew) * (gudtObject(lngObject).udtSystems.sngEnergy / gudtGenerator(gudtObject(lngObject).udtSystems.bytGenerator).lngMaxBattery)
    
    'Player?
    If blnPlayer = True Then
        'Yep!
        sngShieldPercent = gudtPlayer.udtSystems.sngShieldEnergy / gudtShield(gudtPlayer.udtSystems.bytShield).lngMaxEnergy
    Else
        'Noop..
        sngShieldPercent = 2 * gudtObject(gudtObject(lngObject).udtAI.lngTarget).udtSystems.sngEnergy / gudtGenerator(gudtObject(gudtObject(lngObject).udtAI.lngTarget).udtSystems.bytGenerator).lngMaxBattery
        If sngShieldPercent > 1 Then sngShieldPercent = 1
    End If
    
    'Do we have any shield to display?
    If sngShieldPercent > 0 Then
        If blnPlayer = True Then
            gudtPlayer.udtInfo.blnShieldUp = True
            gudtPlayer.udtInfo.lngShieldDown = glngGameTime + SHIELD_DURATION
        Else
            gudtObject(gudtObject(lngObject).udtAI.lngTarget).udtInfo.blnShieldUp = True
            gudtObject(gudtObject(lngObject).udtAI.lngTarget).udtInfo.lngShieldDown = glngGameTime + SHIELD_DURATION
        End If
    End If
    
    'Calc energy loss
    sngEnergyLoss = sngConcussive * sngShieldPercent
    'Calc damage
    sngArmourLoss = sngConcussive * (1 - sngShieldPercent)
    'Calc crew loss
    sngCrewLoss = sngRadiation * (1 - sngShieldPercent)
    
    'Subtract
    If blnPlayer = True Then
        'Energy
        If gudtPlayer.udtSystems.sngEnergy >= sngEnergyLoss Then
            gudtPlayer.udtSystems.sngEnergy = gudtPlayer.udtSystems.sngEnergy - sngEnergyLoss
        Else
            gudtPlayer.udtSystems.sngShieldEnergy = gudtPlayer.udtSystems.sngShieldEnergy + gudtPlayer.udtSystems.sngEnergy - sngEnergyLoss
        End If
        'Damage
        gudtPlayer.udtSystems.lngArmour = gudtPlayer.udtSystems.lngArmour - sngArmourLoss
        If gudtPlayer.udtSystems.lngArmour < 0 Then blnDead = True
        'Crew
        gudtPlayer.udtSystems.lngCrew = gudtPlayer.udtSystems.lngCrew - sngCrewLoss
        If gudtPlayer.udtSystems.lngCrew < 0 Then gudtPlayer.udtSystems.lngCrew = 0
    Else
        'Energy
        gudtObject(gudtObject(lngObject).udtAI.lngTarget).udtSystems.sngEnergy = gudtObject(gudtObject(lngObject).udtAI.lngTarget).udtSystems.sngEnergy - sngEnergyLoss
        'Damage
        gudtObject(gudtObject(lngObject).udtAI.lngTarget).udtSystems.lngArmour = gudtObject(gudtObject(lngObject).udtAI.lngTarget).udtSystems.lngArmour - sngArmourLoss
        If gudtObject(gudtObject(lngObject).udtAI.lngTarget).udtSystems.lngArmour < 0 Then blnDead = True
        'crew
        gudtObject(gudtObject(lngObject).udtAI.lngTarget).udtSystems.lngCrew = gudtObject(gudtObject(lngObject).udtAI.lngTarget).udtSystems.lngCrew - sngCrewLoss
        If gudtObject(gudtObject(lngObject).udtAI.lngTarget).udtSystems.lngCrew < 0 Then gudtObject(gudtObject(lngObject).udtAI.lngTarget).udtSystems.lngCrew = 0
    End If
    
    'Dead?
    If blnDead Then
        'Player?
        If blnPlayer = True Then
            'Explosion
            CreateExplosion gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, gudtPlayer.udtPhysics.sngSpeed, gudtPlayer.udtPhysics.sngHeading
            'Dead
            PlayerDead
        Else
            'Make an explosion
            CreateExplosion gudtObject(gudtObject(lngObject).udtAI.lngTarget).udtPhysics.dblX, gudtObject(gudtObject(lngObject).udtAI.lngTarget).udtPhysics.dblY, gudtObject(gudtObject(lngObject).udtAI.lngTarget).udtPhysics.sngSpeed, gudtObject(gudtObject(lngObject).udtAI.lngTarget).udtPhysics.sngHeading
            'It's dead
            ObjectDead gudtObject(lngObject).udtAI.lngTarget
        End If
    Else
        'Player?
        If blnPlayer = True Then
            'Create a new laser
            CreateLaser lngObject, TARGET_PLAYER, gudtObject(lngObject).udtSystems.bytLaser
        Else
            'Create a new laser
            CreateLaser lngObject, gudtObject(lngObject).udtAI.lngTarget, gudtObject(lngObject).udtSystems.bytLaser
        End If
        'Lasers ON!
        gudtObject(lngObject).udtAI.blnLaserFire = True
    End If

End Sub

Sub CreateExplosion(dblX As Double, dblY As Double, sngSpeed As Single, sngDirection As Single, Optional bytExplosionType As Byte = 0)

Dim lngPan As Long
Dim lngVolume As Long

    'Make a new spot
    ReDim Preserve gudtExplosion(glngNumExplosions)
    glngNumExplosions = glngNumExplosions + 1
    
    'Enter the data
    gudtExplosion(glngNumExplosions - 1).bytAnimFrame = 0
    gudtExplosion(glngNumExplosions - 1).dblX = dblX
    gudtExplosion(glngNumExplosions - 1).dblY = dblY
    gudtExplosion(glngNumExplosions - 1).lngNextFrameTime = glngGameTime + EXPLOSION_ANIM_RATE
    gudtExplosion(glngNumExplosions - 1).sngDirection = sngDirection
    gudtExplosion(glngNumExplosions - 1).sngSpeed = sngSpeed
    gudtExplosion(glngNumExplosions - 1).bytExplosionType = bytExplosionType
    
    'Get pan + vol
    GetPanAndVol dblX, dblY, lngPan, lngVolume
    'Play the sound
    If bytExplosionType = 0 Then
        DSound.PlaySound glngExplosionSmall, False, True, False, lngPan, lngVolume
    Else
        DSound.PlaySound glngExplosionBig, False, True, False, lngPan, lngVolume
    End If

End Sub

Sub DeleteExplosion(lngIndex As Long)

Dim i As Long

    'Is there such an explosion?
    If lngIndex >= glngNumExplosions Then Exit Sub

    'If this is the last explosion, so be it
    If glngNumExplosions = 1 Then
        Erase gudtExplosion
        glngNumExplosions = 0
        Exit Sub
    End If
    
    'Otherwise, remove and decrement!
    For i = lngIndex To glngNumExplosions - 2
        gudtExplosion(i).bytAnimFrame = gudtExplosion(i + 1).bytAnimFrame
        gudtExplosion(i).dblX = gudtExplosion(i + 1).dblX
        gudtExplosion(i).dblY = gudtExplosion(i + 1).dblY
        gudtExplosion(i).lngNextFrameTime = gudtExplosion(i + 1).lngNextFrameTime
        gudtExplosion(i).sngDirection = gudtExplosion(i + 1).sngDirection
        gudtExplosion(i).sngSpeed = gudtExplosion(i + 1).sngSpeed
        gudtExplosion(i).bytExplosionType = gudtExplosion(i + 1).bytExplosionType
    Next i
    ReDim Preserve gudtExplosion(glngNumExplosions - 2)
    glngNumExplosions = glngNumExplosions - 1

End Sub

Sub UpdateExplosions()

Dim i As Long

    'Loop through the explosions
    i = 0
    Do While i <= glngNumExplosions - 1
        'Move the explosion
        Motion gudtExplosion(i).dblX, gudtExplosion(i).dblY, gudtExplosion(i).sngSpeed, gudtExplosion(i).sngDirection
        'Increment animation?
        If gudtExplosion(i).lngNextFrameTime < glngGameTime Then
            'Are we done?
            If gudtExplosion(i).bytAnimFrame = EXPLOSION_SPRITE_NUM Then
                'Kill the explosion
                DeleteExplosion i
                i = i - 1
            Else
                'Display next frame
                gudtExplosion(i).bytAnimFrame = gudtExplosion(i).bytAnimFrame + 1
                gudtExplosion(i).lngNextFrameTime = glngGameTime + EXPLOSION_ANIM_RATE
            End If
        End If
        'Increment
        i = i + 1
    Loop

End Sub

Private Sub PlayerPhysics()

Dim sngCrew As Single
Dim blnDistribute As Boolean
Dim sngTemp As Single
Dim i As Long

    'Calc crew effectiveness
    sngCrew = gudtPlayer.udtSystems.lngCrew / gudtHull(gudtPlayer.udtSystems.bytHull).lngMaxCrew

    'Set energy levels according to sliders (maximally)
    gudtPlayer.udtSystems.sngGeneratorEnergy = gudtGenerator(gudtPlayer.udtSystems.bytGenerator).lngMaxEnergy * gudtPlayer.udtControl.bytGenerator / BAR_WIDTH
    If gudtPlayer.udtSystems.sngEngineEnergy > gudtEngine(gudtPlayer.udtSystems.bytEngine).lngMaxEnergy * gudtPlayer.udtControl.bytEngine / BAR_WIDTH Then gudtPlayer.udtSystems.sngEngineEnergy = gudtEngine(gudtPlayer.udtSystems.bytEngine).lngMaxEnergy * gudtPlayer.udtControl.bytEngine / BAR_WIDTH
    If gudtPlayer.udtSystems.sngShieldEnergy > gudtShield(gudtPlayer.udtSystems.bytShield).lngMaxEnergy * gudtPlayer.udtControl.bytShield / BAR_WIDTH Then gudtPlayer.udtSystems.sngShieldEnergy = gudtShield(gudtPlayer.udtSystems.bytShield).lngMaxEnergy * gudtPlayer.udtControl.bytShield / BAR_WIDTH
    If gudtPlayer.udtSystems.sngWeaponEnergy > (gudtCannon(gudtPlayer.udtSystems.bytCannon).lngMaxEnergy + gudtLaser(gudtPlayer.udtSystems.bytLaser).lngMaxEnergy) * gudtPlayer.udtControl.bytWeapons / BAR_WIDTH Then gudtPlayer.udtSystems.sngWeaponEnergy = (gudtCannon(gudtPlayer.udtSystems.bytCannon).lngMaxEnergy + gudtLaser(gudtPlayer.udtSystems.bytLaser).lngMaxEnergy) * gudtPlayer.udtControl.bytWeapons / BAR_WIDTH
    'Consume fuel and create energy
    If gudtPlayer.udtSystems.sngGeneratorEnergy > 0 And gudtPlayer.udtSystems.sngFuel > 0 Then
        gudtPlayer.udtSystems.sngFuel = gudtPlayer.udtSystems.sngFuel - gudtGenerator(gudtPlayer.udtSystems.bytGenerator).sngConsumption * (gudtPlayer.udtSystems.sngGeneratorEnergy / gudtGenerator(gudtPlayer.udtSystems.bytGenerator).lngMaxEnergy) * glngElapsed * sngCrew
        gudtPlayer.udtSystems.sngEnergy = gudtPlayer.udtSystems.sngEnergy + gudtGenerator(gudtPlayer.udtSystems.bytGenerator).sngOutPut * (gudtPlayer.udtSystems.sngGeneratorEnergy / gudtGenerator(gudtPlayer.udtSystems.bytGenerator).lngMaxEnergy) * glngElapsed * sngCrew
        'Ensure we're not over the max energy
        If gudtPlayer.udtSystems.sngEnergy > gudtGenerator(gudtPlayer.udtSystems.bytGenerator).lngMaxBattery Then gudtPlayer.udtSystems.sngEnergy = gudtGenerator(gudtPlayer.udtSystems.bytGenerator).lngMaxBattery
        'Ensure we're not below zero fuel
        If gudtPlayer.udtSystems.sngFuel < 0 Then gudtPlayer.udtSystems.sngFuel = 0
    End If
    'Jammer
    If gudtPlayer.udtSystems.blnJammerActive = True Then gudtPlayer.udtSystems.sngEnergy = gudtPlayer.udtSystems.sngEnergy - gudtJammer.sngConsumption * glngElapsed
    If gudtPlayer.udtSystems.sngEnergy <= 0 Then gudtPlayer.udtSystems.blnJammerActive = False
    'Consume energy
    If gudtPlayer.udtSystems.sngEngineEnergy > 0 Then gudtPlayer.udtSystems.sngEnergy = gudtPlayer.udtSystems.sngEnergy - gudtEngine(gudtPlayer.udtSystems.bytEngine).sngConsumption * gudtPlayer.udtSystems.sngEngineEnergy / gudtEngine(gudtPlayer.udtSystems.bytEngine).lngMaxEnergy * glngElapsed
    If gudtPlayer.udtSystems.sngShieldEnergy > 0 Then gudtPlayer.udtSystems.sngEnergy = gudtPlayer.udtSystems.sngEnergy - gudtShield(gudtPlayer.udtSystems.bytShield).sngConsumption * gudtPlayer.udtSystems.sngShieldEnergy / gudtShield(gudtPlayer.udtSystems.bytShield).lngMaxEnergy * glngElapsed
    If gudtPlayer.udtSystems.sngWeaponEnergy > 0 Then gudtPlayer.udtSystems.sngEnergy = gudtPlayer.udtSystems.sngEnergy - (gudtCannon(gudtPlayer.udtSystems.bytCannon).sngConsumption + gudtLaser(gudtPlayer.udtSystems.bytLaser).sngConsumption) * gudtPlayer.udtSystems.sngWeaponEnergy / (gudtCannon(gudtPlayer.udtSystems.bytCannon).lngMaxEnergy + gudtLaser(gudtPlayer.udtSystems.bytLaser).lngMaxEnergy) * glngElapsed
    'Distribute (or remove) energy to (from) the 4 systems
    blnDistribute = True
    Do While blnDistribute
        'Store the energy value
        sngTemp = gudtPlayer.udtSystems.sngEnergy
        'Check if we're below the sliders, and if so, add energy!
        If gudtPlayer.udtSystems.sngEnergy > 0.01 And gudtPlayer.udtSystems.sngEngineEnergy < gudtEngine(gudtPlayer.udtSystems.bytEngine).lngMaxEnergy * gudtPlayer.udtControl.bytEngine / BAR_WIDTH Then
            gudtPlayer.udtSystems.sngEngineEnergy = gudtPlayer.udtSystems.sngEngineEnergy + 0.01
            gudtPlayer.udtSystems.sngEnergy = gudtPlayer.udtSystems.sngEnergy - 0.01
        End If
        If gudtPlayer.udtSystems.sngEnergy > 0.01 And gudtPlayer.udtSystems.sngShieldEnergy < gudtShield(gudtPlayer.udtSystems.bytShield).lngMaxEnergy * gudtPlayer.udtControl.bytShield / BAR_WIDTH Then
            gudtPlayer.udtSystems.sngShieldEnergy = gudtPlayer.udtSystems.sngShieldEnergy + 0.01
            gudtPlayer.udtSystems.sngEnergy = gudtPlayer.udtSystems.sngEnergy - 0.01
        End If
        If gudtPlayer.udtSystems.sngEnergy > 0.01 And gudtPlayer.udtSystems.sngWeaponEnergy < (gudtCannon(gudtPlayer.udtSystems.bytCannon).lngMaxEnergy + gudtLaser(gudtPlayer.udtSystems.bytLaser).lngMaxEnergy) * gudtPlayer.udtControl.bytWeapons / BAR_WIDTH Then
            gudtPlayer.udtSystems.sngWeaponEnergy = gudtPlayer.udtSystems.sngWeaponEnergy + 0.01
            gudtPlayer.udtSystems.sngEnergy = gudtPlayer.udtSystems.sngEnergy - 0.01
        End If
        'Check if we've got negative energy, and if so, remove it from the systems!
        If gudtPlayer.udtSystems.sngEnergy < 0 And gudtPlayer.udtSystems.sngEngineEnergy > 0 Then
            gudtPlayer.udtSystems.sngEnergy = gudtPlayer.udtSystems.sngEnergy + 0.01
            gudtPlayer.udtSystems.sngEngineEnergy = gudtPlayer.udtSystems.sngEngineEnergy - 0.01
        End If
        If gudtPlayer.udtSystems.sngEnergy < 0 And gudtPlayer.udtSystems.sngShieldEnergy > 0 Then
            gudtPlayer.udtSystems.sngEnergy = gudtPlayer.udtSystems.sngEnergy + 0.01
            gudtPlayer.udtSystems.sngShieldEnergy = gudtPlayer.udtSystems.sngShieldEnergy - 0.01
        End If
        If gudtPlayer.udtSystems.sngEnergy < 0 And gudtPlayer.udtSystems.sngWeaponEnergy > 0 Then
            gudtPlayer.udtSystems.sngEnergy = gudtPlayer.udtSystems.sngEnergy + 0.01
            gudtPlayer.udtSystems.sngWeaponEnergy = gudtPlayer.udtSystems.sngWeaponEnergy - 0.01
        End If
        'Check if there was any change in energy
        If sngTemp = gudtPlayer.udtSystems.sngEnergy Then blnDistribute = False
    Loop
    'Ensure no negative values
    If gudtPlayer.udtSystems.sngEnergy < 0 Then gudtPlayer.udtSystems.sngEnergy = 0
    If gudtPlayer.udtSystems.sngEngineEnergy < 0 Then gudtPlayer.udtSystems.sngEngineEnergy = 0
    If gudtPlayer.udtSystems.sngShieldEnergy < 0 Then gudtPlayer.udtSystems.sngShieldEnergy = 0
    If gudtPlayer.udtSystems.sngWeaponEnergy < 0 Then gudtPlayer.udtSystems.sngWeaponEnergy = 0
    
    'Animate
    If gudtPlayer.udtSprite.bytAnimAmt > 0 Then
        gudtPlayer.udtSprite.lngAnimLast = gudtPlayer.udtSprite.lngAnimLast + glngElapsed
        If gudtPlayer.udtSprite.lngAnimLast > gudtPlayer.udtSprite.lngAnimRate Then
            'Do it!
            gudtPlayer.udtSprite.lngAnimLast = 0
            gudtPlayer.udtSprite.bytAnimNum = gudtPlayer.udtSprite.bytAnimNum + 1
            If gudtPlayer.udtSprite.bytAnimNum > gudtPlayer.udtSprite.bytAnimAmt Then gudtPlayer.udtSprite.bytAnimNum = 0
        End If
    End If

    'Turning
    If gudtPlayer.udtPhysics.blnTurningRight Then gudtPlayer.udtPhysics.sngFacing = FixAngle(gudtPlayer.udtPhysics.sngFacing + gudtPlayer.udtSystems.sngRotationRate * glngElapsed * sngCrew)
    If gudtPlayer.udtPhysics.blnTurningLeft Then gudtPlayer.udtPhysics.sngFacing = FixAngle(gudtPlayer.udtPhysics.sngFacing - gudtPlayer.udtSystems.sngRotationRate * glngElapsed * sngCrew)
    gudtPlayer.udtSprite.bytFrameNum = Fix((gudtPlayer.udtSprite.bytFrameAmt + 1) * (gudtPlayer.udtPhysics.sngFacing / (2 * Pi)))
    If gudtPlayer.udtSprite.bytFrameNum > FRAME_NUM Then gudtPlayer.udtSprite.bytFrameNum = FRAME_NUM
    
    'Calc mass
    gudtPlayer.udtPhysics.lngMass = gudtHull(gudtPlayer.udtSystems.bytHull).lngMass
    'Sum up all cargo
    If gudtPlayer.udtCargo.lngNumCargo > 0 Then
        For i = 0 To gudtPlayer.udtCargo.lngNumCargo - 1
            gudtPlayer.udtPhysics.lngMass = gudtPlayer.udtPhysics.lngMass + gudtPlayer.udtCargo.lngAmount(i)
        Next i
    End If
    'Add salvage
    gudtPlayer.udtPhysics.lngMass = gudtPlayer.udtPhysics.lngMass + gudtPlayer.udtCargo.lngSalvage
    
    'If we're NOT Faster-Than-Light..
    If Not (gudtPlayer.udtSystems.blnARCDActive = True Or gudtPlayer.udtSystems.blnFTLDActive = True) Then
        'Thrusting
        If gudtPlayer.udtPhysics.blnThrusting Then AddVectors gudtPlayer.udtPhysics.sngSpeed, gudtPlayer.udtPhysics.sngHeading, gudtEngine(gudtPlayer.udtSystems.bytEngine).sngThrust * gudtPlayer.udtSystems.lngCrew / gudtHull(gudtPlayer.udtSystems.bytHull).lngMaxCrew * gudtPlayer.udtSystems.sngEngineEnergy / gudtEngine(gudtPlayer.udtSystems.bytEngine).lngMaxEnergy * glngElapsed / gudtPlayer.udtPhysics.lngMass, gudtPlayer.udtPhysics.sngFacing, gudtPlayer.udtPhysics.sngSpeed, gudtPlayer.udtPhysics.sngHeading
        If gudtPlayer.udtPhysics.blnReverseThrusting Then AddVectors gudtPlayer.udtPhysics.sngSpeed, gudtPlayer.udtPhysics.sngHeading, gudtEngine(gudtPlayer.udtSystems.bytEngine).sngThrust * gudtPlayer.udtSystems.lngCrew / gudtHull(gudtPlayer.udtSystems.bytHull).lngMaxCrew * gudtPlayer.udtSystems.sngEngineEnergy / gudtEngine(gudtPlayer.udtSystems.bytEngine).lngMaxEnergy * glngElapsed / gudtPlayer.udtPhysics.lngMass / 2, gudtPlayer.udtPhysics.sngFacing + Pi, gudtPlayer.udtPhysics.sngSpeed, gudtPlayer.udtPhysics.sngHeading
        'Cap speed if FTLD is not engaged
        If gudtPlayer.udtPhysics.sngSpeed > gudtHull(gudtPlayer.udtSystems.bytHull).sngMaxSpeed Then gudtPlayer.udtPhysics.sngSpeed = gudtHull(gudtPlayer.udtSystems.bytHull).sngMaxSpeed
    Else
        'If we ARE FTL, set speed
        If gudtPlayer.udtSystems.blnARCDActive = True Then gudtPlayer.udtPhysics.sngSpeed = ARCD_SPEED
        If gudtPlayer.udtSystems.blnFTLDActive = True Then gudtPlayer.udtPhysics.sngSpeed = FTLD_SPEED
        'Set direction
        gudtPlayer.udtPhysics.sngHeading = gudtPlayer.udtPhysics.sngFacing
    End If
    
    'Motion
    Motion gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, gudtPlayer.udtPhysics.sngSpeed, gudtPlayer.udtPhysics.sngHeading
            
End Sub

Private Sub AI(lngObject As Long)

Dim bytRetval As Byte
Dim sngDesiredFacing As Single
Dim lngTempDist As Long
Dim blnFiringLasers As Boolean
Dim blnThrusting As Boolean
Dim lngPan As Long
Dim lngVolume As Long

    'Ensure the object exists!
    If gudtObject(lngObject).blnExists = False Then Exit Sub

    'Determine action type
    blnFiringLasers = gudtObject(lngObject).udtAI.blnLaserFire
    Select Case gudtObject(lngObject).udtAI.bytAction
        Case AI_SEEK
            'Determine target
            If gudtObject(lngObject).udtAI.lngTarget = TARGET_PLAYER Then
                'Seek player
                bytRetval = SeekTarget(gudtObject(lngObject).udtPhysics.sngSpeed, gudtEngine(gudtObject(lngObject).udtSystems.bytEngine).sngThrust * gudtObject(lngObject).udtSystems.lngCrew / gudtHull(gudtObject(lngObject).udtSystems.bytHull).lngMaxCrew * gudtObject(lngObject).udtSystems.sngEngineEnergy / gudtEngine(gudtObject(lngObject).udtSystems.bytEngine).lngMaxEnergy / gudtObject(lngObject).udtPhysics.lngMass, gudtObject(lngObject).udtPhysics.sngHeading, gudtObject(lngObject).udtPhysics.sngFacing, gudtObject(lngObject).udtPhysics.dblX, gudtObject(lngObject).udtPhysics.dblY, gudtPlayer.udtPhysics.sngSpeed, gudtPlayer.udtPhysics.sngHeading, gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, sngDesiredFacing, gudtObject(lngObject).udtAI.sngMinDist, gudtObject(lngObject).udtAI.sngSeekDist, gudtObject(lngObject).udtAI.sngTargetBias)
            ElseIf gudtObject(lngObject).udtAI.lngTarget = TARGET_COORDS Then
                'Seek coords
                bytRetval = SeekTarget(gudtObject(lngObject).udtPhysics.sngSpeed, gudtEngine(gudtObject(lngObject).udtSystems.bytEngine).sngThrust * gudtObject(lngObject).udtSystems.lngCrew / gudtHull(gudtObject(lngObject).udtSystems.bytHull).lngMaxCrew * gudtObject(lngObject).udtSystems.sngEngineEnergy / gudtEngine(gudtObject(lngObject).udtSystems.bytEngine).lngMaxEnergy / gudtObject(lngObject).udtPhysics.lngMass, gudtObject(lngObject).udtPhysics.sngHeading, gudtObject(lngObject).udtPhysics.sngFacing, gudtObject(lngObject).udtPhysics.dblX, gudtObject(lngObject).udtPhysics.dblY, 0, 0, gudtObject(lngObject).udtAI.dblX, gudtObject(lngObject).udtAI.dblY, sngDesiredFacing, gudtObject(lngObject).udtAI.sngMinDist, gudtObject(lngObject).udtAI.sngSeekDist, gudtObject(lngObject).udtAI.sngTargetBias)
            ElseIf gudtObject(lngObject).udtAI.lngTarget <> TARGET_NONE Then
                'Seek object
                bytRetval = SeekTarget(gudtObject(lngObject).udtPhysics.sngSpeed, gudtEngine(gudtObject(lngObject).udtSystems.bytEngine).sngThrust * gudtObject(lngObject).udtSystems.lngCrew / gudtHull(gudtObject(lngObject).udtSystems.bytHull).lngMaxCrew * gudtObject(lngObject).udtSystems.sngEngineEnergy / gudtEngine(gudtObject(lngObject).udtSystems.bytEngine).lngMaxEnergy / gudtObject(lngObject).udtPhysics.lngMass, gudtObject(lngObject).udtPhysics.sngHeading, gudtObject(lngObject).udtPhysics.sngFacing, gudtObject(lngObject).udtPhysics.dblX, gudtObject(lngObject).udtPhysics.dblY, gudtObject(gudtObject(lngObject).udtAI.lngTarget).udtPhysics.sngSpeed, gudtObject(gudtObject(lngObject).udtAI.lngTarget).udtPhysics.sngHeading, gudtObject(gudtObject(lngObject).udtAI.lngTarget).udtPhysics.dblX, gudtObject(gudtObject(lngObject).udtAI.lngTarget).udtPhysics.dblY, sngDesiredFacing, gudtObject(lngObject).udtAI.sngMinDist, gudtObject(lngObject).udtAI.sngSeekDist, gudtObject(lngObject).udtAI.sngTargetBias)
            End If
        Case AI_ATTACK
            'If we have no target, find one!
            If gudtObject(lngObject).udtAI.lngTarget = TARGET_NONE Then gudtObject(lngObject).udtAI.lngTarget = FindClosestEnemy(lngObject)
            'Is the targetlock time expired?
            If gudtObject(lngObject).udtAI.lngNewTargetTime <= glngGameTime Then
                gudtObject(lngObject).udtAI.lngTarget = FindClosestEnemy(lngObject)
                gudtObject(lngObject).udtAI.lngNewTargetTime = glngGameTime + gudtObject(lngObject).udtAI.lngLengthTargetLock
            End If
            'Determine target.  Player, or some other object?
            If gudtObject(lngObject).udtAI.lngTarget = TARGET_PLAYER Then
                'Seek player
                If gudtObject(lngObject).udtSystems.bytEngine > 0 Then bytRetval = SeekTarget(gudtObject(lngObject).udtPhysics.sngSpeed, gudtEngine(gudtObject(lngObject).udtSystems.bytEngine).sngThrust * gudtObject(lngObject).udtSystems.lngCrew / gudtHull(gudtObject(lngObject).udtSystems.bytHull).lngMaxCrew * gudtObject(lngObject).udtSystems.sngEngineEnergy / gudtEngine(gudtObject(lngObject).udtSystems.bytEngine).lngMaxEnergy / gudtObject(lngObject).udtPhysics.lngMass, gudtObject(lngObject).udtPhysics.sngHeading, gudtObject(lngObject).udtPhysics.sngFacing, gudtObject(lngObject).udtPhysics.dblX, gudtObject(lngObject).udtPhysics.dblY, gudtPlayer.udtPhysics.sngSpeed, gudtPlayer.udtPhysics.sngHeading, gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, sngDesiredFacing, gudtObject(lngObject).udtAI.sngMinDist, gudtObject(lngObject).udtAI.sngSeekDist, gudtObject(lngObject).udtAI.sngTargetBias, gudtCannon(gudtObject(lngObject).udtSystems.bytCannon).sngSpeed)
                'Check dist
                If gudtObject(lngObject).udtInfo.dblDistance < GREAT_DIST Then
                    'Fire cannons?
                    lngTempDist = gudtObject(lngObject).udtInfo.dblDistance
                    If lngTempDist <= gudtObject(lngObject).udtAI.sngCannonDist Then
                        'Is it a turretted cannon, or are we within aim tolerance?
                        If (gudtCannon(gudtObject(lngObject).udtSystems.bytCannon).lngCannonType And CANNON_TURRET = CANNON_TURRET) Or (AngleDifference(FindAngle(gudtObject(lngObject).udtPhysics.dblX, gudtObject(lngObject).udtPhysics.dblY, gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY), gudtObject(lngObject).udtPhysics.sngFacing) <= gudtObject(lngObject).udtAI.sngAimTolerance) Then
                            ObjectFireCannon lngObject
                        End If
                    End If
                    'Fire lasers?
                    If ((gudtObject(lngObject).udtSystems.sngEnergy / gudtGenerator(gudtObject(lngObject).udtSystems.bytGenerator).lngMaxBattery) >= LASER_FIRE_RECHARGE) Then
                        'If we're past the recharge value, always fire
                        ObjectFireLaser lngObject
                    ElseIf ((gudtObject(lngObject).udtSystems.sngEnergy / gudtGenerator(gudtObject(lngObject).udtSystems.bytGenerator).lngMaxBattery) >= LASER_FIRE_MIN) And (gudtObject(lngObject).udtAI.blnLaserFire = True) Then
                        'If we're ALREADY firing, continue until the minimum
                        ObjectFireLaser lngObject
                    Else
                        'Lasers OFF!
                        gudtObject(lngObject).udtAI.blnLaserFire = False
                    End If
                    'Fire missiles?
                    If gudtObject(lngObject).udtInfo.dblDistance < MISSILE_FIRE_DIST Then
                        'Do we have any?
                        If gudtObject(lngObject).udtSystems.intMissileNum > 0 Then
                            'Is it time?
                            If gudtObject(lngObject).udtSystems.lngMissileLastFire + gudtMissile(gudtObject(lngObject).udtSystems.bytMissile).lngFireRate <= glngGameTime Then
                                'Fire away!
                                CreateMissile gudtObject(lngObject).udtSystems.bytMissile, gudtObject(lngObject).udtPhysics.dblX, gudtObject(lngObject).udtPhysics.dblY, lngObject, -1, gudtObject(lngObject).udtPhysics.sngSpeed, gudtObject(lngObject).udtPhysics.sngHeading, gudtObject(lngObject).udtPhysics.sngFacing
                                'Decrement missiles
                                gudtObject(lngObject).udtSystems.intMissileNum = gudtObject(lngObject).udtSystems.intMissileNum - 1
                                'Set fire timer
                                gudtObject(lngObject).udtSystems.lngMissileLastFire = glngGameTime
                            End If
                        End If
                    End If
                End If
            ElseIf gudtObject(lngObject).udtAI.lngTarget <> TARGET_NONE Then
                'Seek object
                If gudtObject(lngObject).udtSystems.bytEngine > 0 Then
                    bytRetval = SeekTarget(gudtObject(lngObject).udtPhysics.sngSpeed, gudtEngine(gudtObject(lngObject).udtSystems.bytEngine).sngThrust * gudtObject(lngObject).udtSystems.lngCrew / gudtHull(gudtObject(lngObject).udtSystems.bytHull).lngMaxCrew * gudtObject(lngObject).udtSystems.sngEngineEnergy / gudtEngine(gudtObject(lngObject).udtSystems.bytEngine).lngMaxEnergy / gudtObject(lngObject).udtPhysics.lngMass, gudtObject(lngObject).udtPhysics.sngHeading, gudtObject(lngObject).udtPhysics.sngFacing, gudtObject(lngObject).udtPhysics.dblX, gudtObject(lngObject).udtPhysics.dblY, gudtObject(gudtObject(lngObject).udtAI.lngTarget).udtPhysics.sngSpeed, gudtObject(gudtObject(lngObject).udtAI.lngTarget).udtPhysics.sngHeading, gudtObject(gudtObject(lngObject).udtAI.lngTarget).udtPhysics.dblX, gudtObject(gudtObject(lngObject).udtAI.lngTarget).udtPhysics.dblY, sngDesiredFacing, gudtObject(lngObject).udtAI.sngMinDist, gudtObject(lngObject).udtAI.sngSeekDist, gudtObject(lngObject).udtAI.sngTargetBias, _
                      gudtCannon(gudtObject(lngObject).udtSystems.bytCannon).sngSpeed)
                End If
                'Check dist
                If GetDist(gudtObject(gudtObject(lngObject).udtAI.lngTarget).udtPhysics.dblX, gudtObject(gudtObject(lngObject).udtAI.lngTarget).udtPhysics.dblY, gudtObject(lngObject).udtPhysics.dblX, gudtObject(lngObject).udtPhysics.dblY) < GREAT_DIST Then
                    'Fire cannons?
                    lngTempDist = GetDist(gudtObject(gudtObject(lngObject).udtAI.lngTarget).udtPhysics.dblX, gudtObject(gudtObject(lngObject).udtAI.lngTarget).udtPhysics.dblY, gudtObject(lngObject).udtPhysics.dblX, gudtObject(lngObject).udtPhysics.dblY)
                    If lngTempDist <= gudtObject(lngObject).udtAI.sngCannonDist Then
                        'Is it a turretted cannon, or are we within aim tolerance?
                        If (gudtCannon(gudtObject(lngObject).udtSystems.bytCannon).lngCannonType And CANNON_TURRET = CANNON_TURRET) Or (AngleDifference(FindAngle(gudtObject(lngObject).udtPhysics.dblX, gudtObject(lngObject).udtPhysics.dblY, gudtObject(gudtObject(lngObject).udtAI.lngTarget).udtPhysics.dblX, gudtObject(gudtObject(lngObject).udtAI.lngTarget).udtPhysics.dblY), gudtObject(lngObject).udtPhysics.sngFacing) <= gudtObject(lngObject).udtAI.sngAimTolerance) Then
                            ObjectFireCannon lngObject
                        End If
                    End If
                    'Fire lasers?
                    If ((gudtObject(lngObject).udtSystems.sngEnergy / gudtGenerator(gudtObject(lngObject).udtSystems.bytGenerator).lngMaxBattery) >= LASER_FIRE_RECHARGE) Then
                        'If we're past the recharge value, always fire
                        ObjectFireLaser lngObject
                    ElseIf ((gudtObject(lngObject).udtSystems.sngEnergy / gudtGenerator(gudtObject(lngObject).udtSystems.bytGenerator).lngMaxBattery) >= LASER_FIRE_MIN) And (gudtObject(lngObject).udtAI.blnLaserFire = True) Then
                        'If we're ALREADY firing, continue until the minimum
                        ObjectFireLaser lngObject
                    Else
                        'Lasers OFF!
                        gudtObject(lngObject).udtAI.blnLaserFire = False
                    End If
                    'Fire missiles?
                    If lngTempDist < MISSILE_FIRE_DIST Then
                        'Do we have any?
                        If gudtObject(lngObject).udtSystems.intMissileNum > 0 Then
                            'Is it time?
                            If gudtObject(lngObject).udtSystems.lngMissileLastFire + gudtMissile(gudtObject(lngObject).udtSystems.bytMissile).lngFireRate <= glngGameTime Then
                                'Fire away!
                                CreateMissile gudtObject(lngObject).udtSystems.bytMissile, gudtObject(lngObject).udtPhysics.dblX, gudtObject(lngObject).udtPhysics.dblY, lngObject, gudtObject(lngObject).udtAI.lngTarget, gudtObject(lngObject).udtPhysics.sngSpeed, gudtObject(lngObject).udtPhysics.sngHeading, gudtObject(lngObject).udtPhysics.sngFacing
                                'Decrement missiles
                                gudtObject(lngObject).udtSystems.intMissileNum = gudtObject(lngObject).udtSystems.intMissileNum - 1
                                'Set fire timer
                                gudtObject(lngObject).udtSystems.lngMissileLastFire = glngGameTime
                            End If
                        End If
                    End If
                End If
            End If
    End Select
    
    'Engine sound stuff
    blnThrusting = gudtObject(lngObject).udtAI.blnThrusting
    gudtObject(lngObject).udtAI.blnThrusting = False
    
    'Take action
    If (bytRetval And ACTION_LEFT) = ACTION_LEFT Then
        gudtObject(lngObject).udtPhysics.blnTurningLeft = True
        gudtObject(lngObject).udtPhysics.blnTurningRight = False
    End If
    If (bytRetval And ACTION_RIGHT) = ACTION_RIGHT Then
        gudtObject(lngObject).udtPhysics.blnTurningLeft = False
        gudtObject(lngObject).udtPhysics.blnTurningRight = True
    End If
    If (bytRetval And ACTION_NOTURN) = ACTION_NOTURN Then
        gudtObject(lngObject).udtPhysics.sngFacing = FixAngle(sngDesiredFacing)
        gudtObject(lngObject).udtPhysics.blnTurningLeft = False
        gudtObject(lngObject).udtPhysics.blnTurningRight = False
    End If
    If (bytRetval And ACTION_THRUST) = ACTION_THRUST Then
        gudtObject(lngObject).udtPhysics.blnThrusting = True
        gudtObject(lngObject).udtPhysics.blnReverseThrusting = False
        gudtObject(lngObject).udtAI.blnThrusting = True
    End If
    If (bytRetval And ACTION_REVERSETHRUST) = ACTION_REVERSETHRUST Then
        gudtObject(lngObject).udtPhysics.blnThrusting = False
        gudtObject(lngObject).udtPhysics.blnReverseThrusting = True
        gudtObject(lngObject).udtAI.blnThrusting = True
    End If
    If (bytRetval And ACTION_NOTHRUST) = ACTION_NOTHRUST Then
        gudtObject(lngObject).udtPhysics.blnThrusting = False
        gudtObject(lngObject).udtPhysics.blnReverseThrusting = False
    End If
    
    'Check for distance
    If gudtObject(lngObject).udtInfo.dblDistance <= MIN_SOUND_DIST Then
        'Get pan + vol
        GetPanAndVol gudtObject(lngObject).udtPhysics.dblX, gudtObject(lngObject).udtPhysics.dblY, lngPan, lngVolume
        'If we WEREN'T firing lasers, and now we are, start sound
        If (blnFiringLasers = False) And (gudtObject(lngObject).udtAI.blnLaserFire = True) Then
            gudtObject(lngObject).udtAI.lngLaserSound = DSound.PlaySound(gudtLaser(gudtObject(lngObject).udtSystems.bytLaser).lngSound, False, True, True, lngPan, lngVolume)
        ElseIf (blnFiringLasers = True) And (gudtObject(lngObject).udtAI.blnLaserFire = False) Then
            'If we WERE firing lasers, and now we're not, stop sound
            DSound.StopSound gudtObject(lngObject).udtAI.lngLaserSound
        ElseIf (blnFiringLasers = True) And (gudtObject(lngObject).udtAI.blnLaserFire = True) Then
            'We're STILL firing.. adjust pan + vol
            DSound.SetVolume gudtObject(lngObject).udtAI.lngLaserSound, lngVolume
            DSound.SetPan gudtObject(lngObject).udtAI.lngLaserSound, lngPan
        End If
'        'If we WEREN'T thrusting, and now we are, start sound
'        If (blnThrusting = False) And (gudtObject(lngObject).udtAI.blnThrusting = True) Then
'            gudtObject(lngObject).udtAI.lngThrustSound = DSound.PlaySound(gudtEngine(gudtObject(lngObject).udtSystems.bytEngine).lngSound, False, True, True, lngPan, lngVolume)
'        ElseIf (blnThrusting = True) And (gudtObject(lngObject).udtAI.blnThrusting = False) Then
'            'If we WRE thrusting, and now we're not, stop sound
'            DSound.StopSound gudtObject(lngObject).udtAI.lngThrustSound
'        ElseIf (blnThrusting = True) And (gudtObject(lngObject).udtAI.blnThrusting = True) Then
'            'If we're STILL thrusting.. adjust pan + vol
'            DSound.SetVolume gudtObject(lngObject).udtAI.lngThrustSound, lngVolume
'            DSound.SetPan gudtObject(lngObject).udtAI.lngThrustSound, lngPan
'        End If
    End If
    
End Sub

Private Sub PlayerAI()

Dim bytRetval As Byte
Dim sngDesiredFacing As Single
Dim blnThrusting As Boolean

    'Determine action type
    Select Case gudtPlayer.udtAI.bytAction
        Case AI_NONE
            'Do nothing!
            Exit Sub
        Case AI_AUTOPILOT
            'Determine target
            If gudtPlayer.udtAI.lngTarget = TARGET_COORDS Then
                'Are we FTL?
                If gudtPlayer.udtSystems.blnARCDActive = True Or gudtPlayer.udtSystems.blnFTLDActive = True Then
                    'Face coords
                     bytRetval = FaceTarget(gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, gudtPlayer.udtAI.dblX, gudtPlayer.udtAI.dblY, gudtPlayer.udtPhysics.sngFacing, sngDesiredFacing)
                     'If we're close, drop out of light
                     If CDbl(AUTOPILOT_FTL_DIST * NORMALIZE_DISTANCE_AU * NORMALIZE_DISTANCE_LY) > GetDist(gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, gudtPlayer.udtAI.dblX, gudtPlayer.udtAI.dblY) Then
                        'Drop out of light
                        gudtPlayer.udtSystems.blnARCDActive = False
                        gudtPlayer.udtSystems.blnFTLDActive = False
                        'Move us close (yes, this is fudging it a bit)
                        If AUTOPILOT_FTL_SETDIST * NORMALIZE_DISTANCE_AU < GetDist(gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, gudtPlayer.udtAI.dblX, gudtPlayer.udtAI.dblY) Then
                            gudtPlayer.udtPhysics.dblX = Sin(-sngDesiredFacing) * AUTOPILOT_FTL_SETDIST * NORMALIZE_DISTANCE_AU + gudtPlayer.udtAI.dblX
                            gudtPlayer.udtPhysics.dblY = Cos(sngDesiredFacing) * AUTOPILOT_FTL_SETDIST * NORMALIZE_DISTANCE_AU + gudtPlayer.udtAI.dblY
                        End If
                    End If
                Else
                    'Seek coords
                    bytRetval = SeekTargetNoRev(gudtPlayer.udtPhysics.sngSpeed, gudtEngine(gudtPlayer.udtSystems.bytEngine).sngThrust * gudtPlayer.udtSystems.lngCrew / gudtHull(gudtPlayer.udtSystems.bytHull).lngMaxCrew * gudtPlayer.udtSystems.sngEngineEnergy / gudtEngine(gudtPlayer.udtSystems.bytEngine).lngMaxEnergy / gudtPlayer.udtPhysics.lngMass, gudtPlayer.udtPhysics.sngHeading, gudtPlayer.udtPhysics.sngFacing, gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, 0, 0, gudtPlayer.udtAI.dblX, gudtPlayer.udtAI.dblY, sngDesiredFacing)
                    'If we're close, stop
                    If AUTOPILOT_DIST > GetDist(gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, gudtPlayer.udtAI.dblX, gudtPlayer.udtAI.dblY) Then
                        gudtPlayer.udtAI.bytAction = AI_ALLSTOP
                        gudtPlayer.udtAI.lngTarget = TARGET_NONE
                    End If
                End If
            ElseIf gudtPlayer.udtAI.lngTarget <> TARGET_NONE Then
                'Are we FTL?
                If gudtPlayer.udtSystems.blnARCDActive = True Or gudtPlayer.udtSystems.blnFTLDActive = True Then
                    'Face object
                    bytRetval = FaceTarget(gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, gudtObject(gudtPlayer.udtAI.lngTarget).udtPhysics.dblX, gudtObject(gudtPlayer.udtAI.lngTarget).udtPhysics.dblY, gudtPlayer.udtPhysics.sngFacing, sngDesiredFacing)
                    'If we're close, drop out of light
                     If CDbl(AUTOPILOT_FTL_DIST * NORMALIZE_DISTANCE_AU * NORMALIZE_DISTANCE_LY) > GetDist(gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, gudtObject(gudtPlayer.udtAI.lngTarget).udtPhysics.dblX, gudtObject(gudtPlayer.udtAI.lngTarget).udtPhysics.dblY) Then
                        'Drop out of light
                        gudtPlayer.udtSystems.blnARCDActive = False
                        gudtPlayer.udtSystems.blnFTLDActive = False
                        'Move us close (yes, this is fudging it a bit)
                        If AUTOPILOT_FTL_SETDIST * NORMALIZE_DISTANCE_AU < GetDist(gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, gudtObject(gudtPlayer.udtAI.lngTarget).udtPhysics.dblX, gudtObject(gudtPlayer.udtAI.lngTarget).udtPhysics.dblY) Then
                            gudtPlayer.udtPhysics.dblX = Sin(-sngDesiredFacing) * AUTOPILOT_FTL_SETDIST * NORMALIZE_DISTANCE_AU + gudtObject(gudtPlayer.udtAI.lngTarget).udtPhysics.dblX
                            gudtPlayer.udtPhysics.dblY = Cos(sngDesiredFacing) * AUTOPILOT_FTL_SETDIST * NORMALIZE_DISTANCE_AU + gudtObject(gudtPlayer.udtAI.lngTarget).udtPhysics.dblY
                        End If
                    End If
                Else
                    'Seek object
                    bytRetval = SeekTargetNoRev(gudtPlayer.udtPhysics.sngSpeed, gudtEngine(gudtPlayer.udtSystems.bytEngine).sngThrust * gudtPlayer.udtSystems.lngCrew / gudtHull(gudtPlayer.udtSystems.bytHull).lngMaxCrew * gudtPlayer.udtSystems.sngEngineEnergy / gudtEngine(gudtPlayer.udtSystems.bytEngine).lngMaxEnergy / gudtPlayer.udtPhysics.lngMass, gudtPlayer.udtPhysics.sngHeading, gudtPlayer.udtPhysics.sngFacing, gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, gudtObject(gudtPlayer.udtAI.lngTarget).udtPhysics.sngSpeed, gudtObject(gudtPlayer.udtAI.lngTarget).udtPhysics.sngHeading, gudtObject(gudtPlayer.udtAI.lngTarget).udtPhysics.dblX, gudtObject(gudtPlayer.udtAI.lngTarget).udtPhysics.dblY, sngDesiredFacing)
                    'If we're close, stop
                    If AUTOPILOT_DIST > GetDist(gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, gudtObject(gudtPlayer.udtAI.lngTarget).udtPhysics.dblX, gudtObject(gudtPlayer.udtAI.lngTarget).udtPhysics.dblY) Then
                        gudtPlayer.udtAI.bytAction = AI_ALLSTOP
                        gudtPlayer.udtAI.lngTarget = TARGET_NONE
                    End If
                End If
            End If
        Case AI_ALLSTOP
            'Drop out of light
            gudtPlayer.udtSystems.blnARCDActive = False
            gudtPlayer.udtSystems.blnFTLDActive = False
            'Stop the ship!
            If gudtPlayer.udtPhysics.sngSpeed > ALLSTOP_SPEED Then
                bytRetval = SeekTargetNoRev(gudtPlayer.udtPhysics.sngSpeed, gudtEngine(gudtPlayer.udtSystems.bytEngine).sngThrust * gudtPlayer.udtSystems.lngCrew / gudtHull(gudtPlayer.udtSystems.bytHull).lngMaxCrew * gudtPlayer.udtSystems.sngEngineEnergy / gudtEngine(gudtPlayer.udtSystems.bytEngine).lngMaxEnergy / gudtPlayer.udtPhysics.lngMass, gudtPlayer.udtPhysics.sngHeading, gudtPlayer.udtPhysics.sngFacing, gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, 0, 0, gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY + 1, sngDesiredFacing)
            Else
                gudtPlayer.udtPhysics.sngSpeed = 0
                gudtPlayer.udtAI.bytAction = AI_NONE
                DSound.StopSound gudtPlayer.udtAI.lngThrustSound
            End If
    End Select

    'Sound stuff, skip if autopilot's not on
    If (gudtPlayer.udtAI.bytAction = AI_AUTOPILOT) Or (gudtPlayer.udtAI.bytAction = AI_ALLSTOP) Then
        blnThrusting = gudtPlayer.udtAI.blnThrusting
        gudtPlayer.udtAI.blnThrusting = False
    End If

    'Take action
    If (bytRetval And ACTION_LEFT) = ACTION_LEFT Then
        gudtPlayer.udtPhysics.blnTurningLeft = True
        gudtPlayer.udtPhysics.blnTurningRight = False
    End If
    If (bytRetval And ACTION_RIGHT) = ACTION_RIGHT Then
        gudtPlayer.udtPhysics.blnTurningLeft = False
        gudtPlayer.udtPhysics.blnTurningRight = True
    End If
    If (bytRetval And ACTION_NOTURN) = ACTION_NOTURN Then
        gudtPlayer.udtPhysics.sngFacing = FixAngle(sngDesiredFacing)
        gudtPlayer.udtPhysics.blnTurningLeft = False
        gudtPlayer.udtPhysics.blnTurningRight = False
    End If
    If (bytRetval And ACTION_THRUST) = ACTION_THRUST Then
        gudtPlayer.udtPhysics.blnThrusting = True
        gudtPlayer.udtPhysics.blnReverseThrusting = False
        gudtPlayer.udtAI.blnThrusting = True
    End If
    If (bytRetval And ACTION_REVERSETHRUST) = ACTION_REVERSETHRUST Then
        gudtPlayer.udtPhysics.blnThrusting = False
        gudtPlayer.udtPhysics.blnReverseThrusting = True
        gudtPlayer.udtAI.blnThrusting = True
    End If
    If (bytRetval = ACTION_NOTHRUST) Then
        gudtPlayer.udtPhysics.blnThrusting = False
        gudtPlayer.udtPhysics.blnReverseThrusting = False
    End If

    'Sound stuff, skip if autopilot's not on
    If (gudtPlayer.udtAI.bytAction = AI_AUTOPILOT) Or (gudtPlayer.udtAI.bytAction = AI_ALLSTOP) Then
        'If we WEREN'T thrusting, but are now, start sound
        If (blnThrusting = False) And (gudtPlayer.udtAI.blnThrusting = True) Then
            gudtPlayer.udtAI.lngThrustSound = DSound.PlaySound(gudtEngine(gudtPlayer.udtSystems.bytEngine).lngSound, False, True, True)
        'If we WERE thrusting, but aren't now, end sound
        ElseIf (blnThrusting = True) And (gudtPlayer.udtAI.blnThrusting = False) Then
            DSound.StopSound gudtPlayer.udtAI.lngThrustSound
        End If
    End If

End Sub
